//! Standalone coordinator supervision; one bounded writer and renderer owner.
use serde_json::{Value, json};
use std::os::{fd::AsRawFd, unix::process::CommandExt};
use std::{
    collections::HashMap,
    io::Read,
    path::PathBuf,
    process::{Child, Command, Stdio},
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, Ordering},
        mpsc,
    },
};
use tauri::{AppHandle, Emitter, Manager};
use tokio::sync::oneshot;
type Reply = Result<Value, String>;
/// How a native action answers its coordinator.
pub type NativeReply = Box<dyn FnOnce(Result<Value, String>) + Send>;

/// What a standalone editor's shell supplies to its coordinator.
pub trait Shell: 'static {
    /// The shell crate's own folder (`env!("CARGO_MANIFEST_DIR")`): from it the
    /// checkout is found when the app runs from source.
    const MANIFEST_DIR: &'static str;
    /// The coordinator's entry, relative to the app root.
    const COORDINATOR_ENTRY: &'static str;
    /// The prefix of the editor's environment variables (`XLSX_EDITOR`, `DOCX_EDITOR`).
    const ENV_PREFIX: &'static str;
    /// The prefix of the editor's own screen events.
    const CHANNEL_PREFIX: &'static str;
    /// The editor's own settings for its coordinator.
    fn configure_coordinator(command: &mut Command, root: &std::path::Path, resources: &std::path::Path, packaged: bool);
    /// A window, dialog or system action the coordinator asked for.
    fn dispatch_native(app: &AppHandle, request: &Value, reply: NativeReply);
    /// The coordinator approved quitting.
    fn quit_approved(app: &AppHandle);
    /// No quit request is pending any more (cancelled, or the coordinator failed).
    fn quit_settled(app: &AppHandle);
}
const MAX_FRAME: usize = 16 * 1024 * 1024;
pub struct Backend {
    input: super::pipe_writer::PipeWriter,
    child: Mutex<Child>,
    pending: Mutex<HashMap<String, oneshot::Sender<Reply>>>,
    next: AtomicU64,
    available: AtomicBool,
    stopped: AtomicBool,
    events: Mutex<super::event_budget::EventBudget>,
    renderer: Mutex<Option<String>>,
    quit_settled: fn(&AppHandle),
}
impl Backend {
    pub fn start<S: Shell>(app: &AppHandle, profile: &super::profile::Profile) -> Result<Arc<Self>, String> {
        let standalone = PathBuf::from(S::MANIFEST_DIR)
            .parent()
            .unwrap()
            .to_path_buf();
        let resources = app.path().resource_dir().map_err(|e| e.to_string())?;
        // A damaged release bundle must fail closed, never run checkout code.
        let packaged = !cfg!(debug_assertions)
            || resources
                .join(format!("app/{}", S::COORDINATOR_ENTRY))
                .is_file();
        let root = if packaged {
            resources.join("app")
        } else {
            standalone.parent().unwrap().to_path_buf()
        };
        let runtime = if packaged {
            resources.join("backend/bun")
        } else {
            root.join(".build/runtimes/bun/darwin-arm64/payload/bun")
        };
        let runtime = runtime
            .canonicalize()
            .map_err(|error| format!("unavailable: Standalone Bun runtime is missing: {error}"))?;
        if packaged
            && (!runtime.is_file()
                || !runtime.starts_with(
                    resources
                        .canonicalize()
                        .map_err(|error| error.to_string())?,
                ))
        {
            return Err("permission_denied: Standalone runtime escapes its bundle".into());
        }
        let lock = profile.duplicate_lock()?;
        let fd = lock.as_raw_fd();
        let mut command = Command::new(runtime);
        command
            .arg("--no-install")
            .arg(root.join(S::COORDINATOR_ENTRY))
            .current_dir(&root)
            // Each editor's own names (XLSX_EDITOR_*, DOCX_EDITOR_*), from the including host.
            .env(format!("{}_USER_DATA", S::ENV_PREFIX), &profile.root)
            .env(format!("{}_PROFILE_LOCK_FD", S::ENV_PREFIX), fd.to_string())
            .env(format!("{}_APP_ROOT", S::ENV_PREFIX), &root)
            .env(format!("{}_RESOURCES", S::ENV_PREFIX), &resources)
            .env(format!("{}_PACKAGED", S::ENV_PREFIX), if packaged { "1" } else { "0" })
            .env_remove("BUN_BE_BUN")
            .env_remove("ELECTRON_RUN_AS_NODE")
            .env_remove("NODE_OPTIONS")
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit())
            .process_group(0);
        S::configure_coordinator(&mut command, &root, &resources, packaged);
        unsafe {
            command.pre_exec(move || {
                let flags = libc::fcntl(fd, libc::F_GETFD);
                if flags < 0 || libc::fcntl(fd, libc::F_SETFD, flags & !libc::FD_CLOEXEC) < 0 {
                    return Err(std::io::Error::last_os_error());
                }
                Ok(())
            });
        }
        let mut child = command
            .spawn()
            .map_err(|e| format!("unavailable: Cannot start standalone coordinator: {e}"))?;
        drop(lock);
        let input = child
            .stdin
            .take()
            .ok_or("unavailable: Coordinator stdin is missing")?;
        let mut output = child
            .stdout
            .take()
            .ok_or("unavailable: Coordinator stdout is missing")?;
        let (input, worker) = super::pipe_writer::PipeWriter::prepare(input, Default::default())?;
        let backend = Arc::new(Self {
            input,
            child: Mutex::new(child),
            pending: Default::default(),
            next: AtomicU64::new(0),
            available: AtomicBool::new(true),
            stopped: AtomicBool::new(false),
            events: Default::default(),
            renderer: Default::default(),
            quit_settled: S::quit_settled,
        });
        let weak = Arc::downgrade(&backend);
        let handle = app.clone();
        worker.spawn(move |error| {
            if let Some(backend) = weak.upgrade() {
                backend.fail(&handle, error);
                backend.stop()
            }
        });
        let (actions, requests) = mpsc::sync_channel::<Value>(32);
        let weak = Arc::downgrade(&backend);
        let handle = app.clone();
        std::thread::spawn(move || {
            while let Ok(request) = requests.recv() {
                let id = request["id"].clone();
                let weak = weak.clone();
                let callback_handle = handle.clone();
                S::dispatch_native(
                    &handle,
                    &request,
                    Box::new(move |result| {
                        if let Some(backend) = weak.upgrade() {
                            let message = match result {
                                Ok(result) => {
                                    json!({"kind":"native-result","id":id,"result":result})
                                }
                                Err(error) => json!({"kind":"native-result","id":id,"error":error}),
                            };
                            if let Err(error) =
                                backend.write(&message, super::pipe_writer::Lane::Native)
                            {
                                backend.fail(&callback_handle, error);
                                backend.stop()
                            }
                        }
                    }),
                );
            }
        });
        let weak = Arc::downgrade(&backend);
        let handle = app.clone();
        std::thread::spawn(move || {
            let result = (|| -> Result<(), String> {
                loop {
                    let mut header = [0u8; 4];
                    output
                        .read_exact(&mut header)
                        .map_err(|e| format!("unavailable: Standalone coordinator stopped: {e}"))?;
                    let size = u32::from_le_bytes(header) as usize;
                    if size == 0 || size > MAX_FRAME {
                        return Err("validation: Invalid standalone frame".into());
                    }
                    let mut bytes = vec![0; size];
                    output
                        .read_exact(&mut bytes)
                        .map_err(|e| format!("unavailable: Incomplete standalone reply: {e}"))?;
                    let mut message: Value = serde_json::from_slice(&bytes)
                        .map_err(|_| "validation: Invalid standalone reply")?;
                    let Some(backend) = weak.upgrade() else {
                        return Ok(());
                    };
                    match message["kind"].as_str() {
                        Some("event") => match message["channel"].as_str() {
                            Some("standalone:exit-approved") => {
                                S::quit_approved(&handle);
                                handle.exit(0);
                            }
                            Some("standalone:exit-cancelled") => S::quit_settled(&handle),
                            Some("standalone:exit-failed") => {
                                backend.fail(
                                    &handle,
                                    "unavailable: Standalone editor initialization failed".into(),
                                );
                            }
                            Some(channel) if channel.starts_with(S::CHANNEL_PREFIX) => {
                                let receipt = backend.events.lock().unwrap().reserve(size)?;
                                message["receipt"] = receipt.into();
                                handle
                                    .emit_to("main", "backend:event", message)
                                    .map_err(|e| {
                                        format!(
                                            "unavailable: Standalone event delivery failed: {e}"
                                        )
                                    })?;
                            }
                            _ => return Err("validation: Unexpected standalone event".into()),
                        },
                        Some("native") => {
                            if let Err(error) = actions.try_send(message) {
                                let (message, detail) = match error {
                                    mpsc::TrySendError::Full(message) => {
                                        (message, "busy: Native queue is full")
                                    }
                                    mpsc::TrySendError::Disconnected(message) => {
                                        (message, "unavailable: Native queue stopped")
                                    }
                                };
                                backend.write(&json!({"kind":"native-result","id":message["id"],"error":detail}),super::pipe_writer::Lane::Native)?;
                            }
                        }
                        _ => {
                            if let Some(id) = message["id"].as_str() {
                                if let Some(reply) = backend.pending.lock().unwrap().remove(id) {
                                    let _ = reply.send(match message["error"].as_str() {
                                        Some(error) => Err(error.into()),
                                        None => Ok(message["result"].clone()),
                                    });
                                }
                            } else {
                                return Err("validation: Reply has no identity".into());
                            }
                        }
                    }
                }
            })();
            if let Err(error) = result {
                if let Some(backend) = weak.upgrade() {
                    backend.fail(&handle, error);
                    backend.stop()
                }
            }
        });
        Ok(backend)
    }
    fn write(&self, value: &Value, lane: super::pipe_writer::Lane) -> Result<(), String> {
        let body = serde_json::to_vec(value).map_err(|e| e.to_string())?;
        if body.is_empty() || body.len() > MAX_FRAME {
            return Err("validation: Standalone request exceeds frame limit".into());
        }
        let mut packet = Vec::with_capacity(body.len() + 4);
        packet.extend_from_slice(&(body.len() as u32).to_le_bytes());
        packet.extend_from_slice(&body);
        self.input.submit(packet, lane)
    }
    pub async fn invoke(&self, channel: &str, args: Vec<Value>) -> Reply {
        let id = self.next.fetch_add(1, Ordering::Relaxed).to_string();
        let (sender, receiver) = oneshot::channel();
        {
            let mut pending = self.pending.lock().unwrap();
            if !self.available.load(Ordering::Acquire) {
                return Err("unavailable: Standalone coordinator has stopped".into());
            }
            if pending.len() >= 512 {
                return Err("busy: Too many standalone requests".into());
            }
            pending.insert(id.clone(), sender);
        }
        let _pending = Pending {
            owner: self,
            id: id.clone(),
        };
        self.write(
            &json!({"kind":"invoke","id":id,"channel":channel,"args":args}),
            super::pipe_writer::Lane::Request,
        )?;
        match tokio::time::timeout(std::time::Duration::from_secs(180), receiver).await {
            Ok(Ok(reply)) => reply,
            Ok(Err(_)) => Err("unavailable: Standalone response closed".into()),
            Err(_) => Err("timeout: Operation outcome is unknown; refresh before retrying".into()),
        }
    }
    pub fn attach(&self, token: String) -> Result<(), String> {
        if token.is_empty() || token.len() > 100 {
            return Err("validation: Invalid standalone renderer identity".into());
        }
        *self.renderer.lock().unwrap() = Some(token);
        Ok(())
    }
    pub fn owns(&self, token: &str) -> Result<(), String> {
        if self.renderer.lock().unwrap().as_deref() != Some(token) {
            return Err("permission_denied: Standalone renderer was retired".into());
        }
        Ok(())
    }
    pub fn acknowledge(&self, receipt: u64) -> Result<(), String> {
        self.events.lock().unwrap().acknowledge(receipt)
    }
    fn fail(&self, app: &AppHandle, error: String) {
        let first = self.available.swap(false, Ordering::AcqRel);
        self.input.close(error.clone());
        for (_, reply) in self.pending.lock().unwrap().drain() {
            let _ = reply.send(Err(error.clone()));
        }
        if first {
            (self.quit_settled)(app);
            let _ = app.emit_to(
                "main",
                "backend:event",
                json!({"channel":"standalone:unavailable","args":[error]}),
            );
        }
    }
    pub fn available(&self) -> bool {
        self.available.load(Ordering::Acquire)
    }
    pub fn stop(&self) {
        if self.stopped.swap(true, Ordering::AcqRel) {
            return;
        }
        self.available.store(false, Ordering::Release);
        self.input
            .close("unavailable: Standalone coordinator stopped".into());
        for (_, reply) in self.pending.lock().unwrap().drain() {
            let _ = reply.send(Err("unavailable: Standalone coordinator stopped".into()));
        }
        if let Ok(mut child) = self.child.lock() {
            let group = -(child.id() as i32);
            unsafe { libc::kill(group, libc::SIGTERM) };
            let start = std::time::Instant::now();
            while start.elapsed() < std::time::Duration::from_secs(2) {
                if unsafe { libc::kill(group, 0) } != 0 {
                    break;
                }
                std::thread::sleep(std::time::Duration::from_millis(25))
            }
            unsafe { libc::kill(group, libc::SIGKILL) };
            let _ = child.wait();
        }
    }
}
struct Pending<'a> {
    owner: &'a Backend,
    id: String,
}
impl Drop for Pending<'_> {
    fn drop(&mut self) {
        self.owner.pending.lock().unwrap().remove(&self.id);
    }
}
