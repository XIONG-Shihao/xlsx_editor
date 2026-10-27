//! A host-written marker interrupts long native work without waiting for its
//! single-threaded request queue. Accepted source/snapshot files are immutable.
use crate::Result;
use std::{
    cell::RefCell,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};
thread_local! {static TOKEN:RefCell<Option<(PathBuf,Option<PathBuf>,Instant,bool)>>=const {RefCell::new(None)};}
pub fn begin(root: &Path) {
    TOKEN.with(|t| {
        *t.borrow_mut() = Some((
            root.join("cancelled"),
            None,
            Instant::now() - Duration::from_secs(1),
            false,
        ))
    });
}
pub(crate) fn preview(root: &Path, request: &str) {
    begin(root);
    TOKEN.with(|token| {
        if let Some((_, marker, _, _)) = token.borrow_mut().as_mut() {
            *marker = Some(root.join(format!("preview-cancel-{request}")));
        }
    });
}
pub fn interrupted() -> bool {
    TOKEN.with(|t| {
        let mut t = t.borrow_mut();
        let Some((file, request, last, stop)) = t.as_mut() else {
            return false;
        };
        if !*stop && last.elapsed() >= Duration::from_millis(10) {
            *stop = file.exists() || request.as_ref().is_some_and(|path| path.exists());
            *last = Instant::now();
        }
        *stop
    })
}
pub fn check() -> Result<()> {
    if interrupted() {
        Err(crate::coded(
            "operation-cancelled",
            "Workbook operation cancelled",
        ))
    } else {
        Ok(())
    }
}

/// Check cancellation even while ZIP streams are inside a library copy loop.
pub struct Reader<R>(pub R);
impl<R: std::io::Read> std::io::Read for Reader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        check().map_err(std::io::Error::other)?;
        let size = buffer.len().min(65536);
        self.0.read(&mut buffer[..size])
    }
}
impl<R: std::io::Seek> std::io::Seek for Reader<R> {
    fn seek(&mut self, position: std::io::SeekFrom) -> std::io::Result<u64> {
        check().map_err(std::io::Error::other)?;
        self.0.seek(position)
    }
}
