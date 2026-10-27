//! Taking a lock without crashing. A thread that panics while holding a lock
//! leaves it "poisoned"; `lock().unwrap()` would then panic in every later
//! caller, so one failed request would take the whole process down with it.
//! `locked()` takes the data as it was left and carries on.
use std::sync::{Mutex, MutexGuard, PoisonError};

pub trait Locked<T> {
    fn locked(&self) -> MutexGuard<'_, T>;
}
impl<T> Locked<T> for Mutex<T> {
    fn locked(&self) -> MutexGuard<'_, T> {
        self.lock().unwrap_or_else(PoisonError::into_inner)
    }
}
