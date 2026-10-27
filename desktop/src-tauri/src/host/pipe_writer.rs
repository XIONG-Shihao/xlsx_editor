//! One owner of the coordinator stdin. Producers never wait for pipe capacity.
//! Budgets include the packet being written, and a stalled partial frame closes
//! the connection: retrying on that stream would corrupt operation attribution.
use super::locked::Locked;
use std::{
    io::Write,
    os::fd::AsRawFd,
    sync::{
        Arc, Mutex,
        mpsc::{self, Receiver, SyncSender},
    },
    time::{Duration, Instant},
};
#[derive(Clone, Copy)]
pub enum Lane {
    Request,
    Native,
}
#[derive(Clone, Copy)]
pub struct Limits {
    pub bytes: usize,
    pub frames: usize,
    pub native_bytes: usize,
    pub native_frames: usize,
    pub deadline: Duration,
}
impl Default for Limits {
    fn default() -> Self {
        Self {
            bytes: 64 * 1024 * 1024,
            frames: 512,
            native_bytes: 32 * 1024 * 1024,
            native_frames: 32,
            deadline: Duration::from_secs(15),
        }
    }
}
#[derive(Default)]
struct Budget {
    bytes: usize,
    frames: usize,
    native_bytes: usize,
    native_frames: usize,
    closed: Option<String>,
}
struct Shared {
    budget: Mutex<Budget>,
    limits: Limits,
}
impl Shared {
    fn close(&self, error: String) -> bool {
        let mut budget = self.budget.locked();
        if budget.closed.is_some() {
            false
        } else {
            budget.closed = Some(error);
            true
        }
    }
    fn error(&self) -> Option<String> {
        self.budget.locked().closed.clone()
    }
}
struct Packet {
    bytes: Vec<u8>,
    lane: Lane,
    shared: Arc<Shared>,
}
impl Drop for Packet {
    fn drop(&mut self) {
        let mut budget = self.shared.budget.locked();
        budget.bytes -= self.bytes.len();
        budget.frames -= 1;
        if matches!(self.lane, Lane::Native) {
            budget.native_bytes -= self.bytes.len();
            budget.native_frames -= 1;
        }
    }
}
pub struct PipeWriter {
    sender: SyncSender<Option<Packet>>,
    shared: Arc<Shared>,
}
pub struct Worker<W> {
    input: W,
    receiver: Receiver<Option<Packet>>,
    shared: Arc<Shared>,
}
impl PipeWriter {
    pub fn prepare<W: Write + AsRawFd + Send + 'static>(
        input: W,
        limits: Limits,
    ) -> Result<(Self, Worker<W>), String> {
        let fd = input.as_raw_fd();
        unsafe {
            let flags = libc::fcntl(fd, libc::F_GETFL);
            if flags < 0 || libc::fcntl(fd, libc::F_SETFL, flags | libc::O_NONBLOCK) < 0 {
                return Err(format!(
                    "unavailable: Cannot supervise coordinator pipe: {}",
                    std::io::Error::last_os_error()
                ));
            }
        }
        let (sender, receiver) = mpsc::sync_channel(limits.frames);
        let shared = Arc::new(Shared {
            budget: Mutex::new(Budget::default()),
            limits,
        });
        Ok((
            Self {
                sender,
                shared: shared.clone(),
            },
            Worker {
                input,
                receiver,
                shared,
            },
        ))
    }
    pub fn submit(&self, bytes: Vec<u8>, lane: Lane) -> Result<(), String> {
        {
            let mut budget = self
                .shared
                .budget
                .lock()
                .map_err(|_| "unavailable: Coordinator writer state failed")?;
            if let Some(error) = &budget.closed {
                return Err(error.clone());
            }
            let limits = self.shared.limits;
            if bytes.len() > limits.bytes.saturating_sub(budget.bytes)
                || budget.frames >= limits.frames
                || matches!(lane, Lane::Native)
                    && (bytes.len() > limits.native_bytes.saturating_sub(budget.native_bytes)
                        || budget.native_frames >= limits.native_frames)
            {
                return Err("busy: Coordinator write queue is full".into());
            }
            budget.bytes += bytes.len();
            budget.frames += 1;
            if matches!(lane, Lane::Native) {
                budget.native_bytes += bytes.len();
                budget.native_frames += 1;
            }
        }
        // Release the lock before enqueue/drop: Packet owns its reservation.
        self.sender
            .try_send(Some(Packet {
                bytes,
                lane,
                shared: self.shared.clone(),
            }))
            .map_err(|_| "unavailable: Coordinator writer stopped".into())
    }
    pub fn close(&self, error: String) {
        self.shared.close(error);
        // Wake an idle worker without a periodic polling loop. If full, an
        // active packet will observe closure and drop the receiver itself.
        let _ = self.sender.try_send(None);
    }
    #[cfg(test)]
    pub fn usage(&self) -> (usize, usize) {
        let budget = self.shared.budget.locked();
        (budget.bytes, budget.frames)
    }
}
impl<W: Write + AsRawFd + Send + 'static> Worker<W> {
    pub fn spawn(
        mut self,
        on_failure: impl Fn(String) + Send + 'static,
    ) -> std::thread::JoinHandle<()> {
        std::thread::spawn(move || {
            while self.shared.error().is_none() {
                match self.receiver.recv() {
                    Ok(Some(packet)) => {
                        if let Err(error) = self.write_packet(&packet.bytes) {
                            if self.shared.close(error.clone()) {
                                on_failure(error);
                            }
                            break;
                        }
                    }
                    Ok(None) | Err(_) => break,
                }
            }
            // Dropping the receiver releases reservations for all unsent work.
        })
    }
    fn write_packet(&mut self, bytes: &[u8]) -> Result<(), String> {
        let started = Instant::now();
        let mut written = 0;
        while written < bytes.len() {
            if let Some(error) = self.shared.error() {
                return Err(error);
            }
            if started.elapsed() >= self.shared.limits.deadline {
                return Err("unavailable: Coordinator pipe stalled; restart required".into());
            }
            match self.input.write(&bytes[written..]) {
                Ok(0) => return Err("unavailable: Coordinator pipe closed".into()),
                Ok(count) => written += count,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                    let mut fd = libc::pollfd {
                        fd: self.input.as_raw_fd(),
                        events: libc::POLLOUT,
                        revents: 0,
                    };
                    let result = unsafe { libc::poll(&mut fd, 1, 25) };
                    if result < 0
                        && std::io::Error::last_os_error().kind() != std::io::ErrorKind::Interrupted
                    {
                        return Err("unavailable: Coordinator pipe polling failed".into());
                    }
                    if fd.revents & (libc::POLLERR | libc::POLLHUP | libc::POLLNVAL) != 0 {
                        return Err("unavailable: Coordinator pipe closed".into());
                    }
                }
                Err(error) => {
                    return Err(format!(
                        "unavailable: Coordinator pipe write failed: {error}"
                    ));
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{
        io::Read,
        os::unix::net::UnixStream,
        sync::atomic::{AtomicUsize, Ordering},
    };
    fn fixture(deadline: Duration) -> (PipeWriter, Worker<UnixStream>, UnixStream) {
        let (input, reader) = UnixStream::pair().unwrap();
        let size = 4096i32;
        assert_eq!(
            unsafe {
                libc::setsockopt(
                    input.as_raw_fd(),
                    libc::SOL_SOCKET,
                    libc::SO_SNDBUF,
                    (&size as *const i32).cast(),
                    std::mem::size_of_val(&size) as _,
                )
            },
            0
        );
        let (writer, worker) = PipeWriter::prepare(
            input,
            Limits {
                bytes: 128 * 1024,
                frames: 4,
                native_bytes: 64 * 1024,
                native_frames: 1,
                deadline,
            },
        )
        .unwrap();
        (writer, worker, reader)
    }
    #[test]
    fn stalled_pipe_admission_is_prompt_counts_active_bytes_and_preserves_order() {
        let (writer, worker, mut reader) = fixture(Duration::from_secs(2));
        let handle = worker.spawn(|error| panic!("{error}"));
        let started = Instant::now();
        writer.submit(vec![1; 65536], Lane::Request).unwrap();
        writer.submit(vec![2; 65536], Lane::Request).unwrap();
        assert!(
            writer
                .submit(vec![3], Lane::Request)
                .unwrap_err()
                .starts_with("busy:")
        );
        assert!(started.elapsed() < Duration::from_millis(100));
        assert_eq!(writer.usage(), (131072, 2));
        let mut bytes = vec![0; 131072];
        reader.read_exact(&mut bytes).unwrap();
        assert!(bytes[..65536].iter().all(|&v| v == 1));
        assert!(bytes[65536..].iter().all(|&v| v == 2));
        writer.close("unavailable: Test finished".into());
        handle.join().unwrap();
        assert_eq!(writer.usage(), (0, 0));
    }
    #[test]
    fn native_budget_is_independent_and_partial_write_shutdown_releases_every_slot() {
        let (writer, worker, _reader) = fixture(Duration::from_secs(2));
        let calls = Arc::new(AtomicUsize::new(0));
        let observed = calls.clone();
        let handle = worker.spawn(move |_| {
            observed.fetch_add(1, Ordering::SeqCst);
        });
        writer.submit(vec![1; 65536], Lane::Native).unwrap();
        assert!(writer.submit(vec![2], Lane::Native).is_err());
        writer.submit(vec![3; 65536], Lane::Request).unwrap();
        let start = Instant::now();
        writer.close("unavailable: Test cancellation".into());
        handle.join().unwrap();
        assert!(start.elapsed() < Duration::from_secs(1));
        assert_eq!(writer.usage(), (0, 0));
        assert_eq!(calls.load(Ordering::SeqCst), 0);
    }
    #[test]
    fn broken_or_stalled_pipe_reports_one_failure_and_rejects_future_packets() {
        for broken in [true, false] {
            let (writer, worker, reader) = fixture(Duration::from_millis(100));
            let kept = if broken {
                drop(reader);
                None
            } else {
                Some(reader)
            };
            let calls = Arc::new(AtomicUsize::new(0));
            let observed = calls.clone();
            let handle = worker.spawn(move |error| {
                assert!(error.starts_with("unavailable:"));
                observed.fetch_add(1, Ordering::SeqCst);
            });
            writer.submit(vec![1; 65536], Lane::Request).unwrap();
            handle.join().unwrap();
            assert_eq!(calls.load(Ordering::SeqCst), 1);
            assert_eq!(writer.usage(), (0, 0));
            assert!(
                writer
                    .submit(vec![2], Lane::Request)
                    .unwrap_err()
                    .starts_with("unavailable:")
            );
            drop(kept);
        }
    }
}
