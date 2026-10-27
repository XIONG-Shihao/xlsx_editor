//! Process and profile supervision for this shell: the coordinator process,
//! the profile lease, memory samples and the pipe writer. Its own copy; the
//! other apps in the repository keep theirs.
pub mod backend;
pub mod event_budget;
pub mod locked;
pub mod memory;
pub mod pipe_writer;
pub mod profile;
