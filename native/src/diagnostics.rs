//! Opt-in, command-local timings. Labels are static: never record workbook data.
//! Disabled commands allocate no trace. Spans are inclusive and may overlap.
use serde_json::{Value, json};
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::time::Instant;

#[derive(Default)]
struct Trace {
    spans: BTreeMap<&'static str, (u64, f64)>,
    counters: BTreeMap<&'static str, u64>,
}
thread_local! {
    static TRACE: RefCell<Option<Trace>> = const { RefCell::new(None) };
}
pub(crate) struct Session {
    previous: Option<Trace>,
    start: Option<Instant>,
}
pub(crate) fn begin(enabled: bool) -> Session {
    Session {
        previous: TRACE.with(|trace| trace.replace(enabled.then(Trace::default))),
        start: enabled.then(Instant::now),
    }
}
impl Session {
    pub(crate) fn finish(&self) -> Option<Value> {
        let start = self.start?;
        TRACE.with(|trace| {
            let trace = trace.borrow();
            let trace = trace.as_ref()?;
            let spans: Vec<Value> = trace
                .spans
                .iter()
                .map(|(name, (calls, ms))| json!({"name":name,"calls":calls,"ms":ms}))
                .collect();
            Some(
                json!({"schema":1,"elapsedMs":start.elapsed().as_secs_f64()*1000.0,
                "spans":spans,"counters":trace.counters,"inclusive":true}),
            )
        })
    }
}
impl Drop for Session {
    fn drop(&mut self) {
        TRACE.with(|trace| {
            trace.replace(self.previous.take());
        });
    }
}
pub(crate) fn enabled() -> bool {
    TRACE.with(|trace| trace.borrow().is_some())
}
pub(crate) struct Span {
    name: &'static str,
    start: Option<Instant>,
}
pub(crate) fn span(name: &'static str) -> Span {
    Span {
        name,
        start: enabled().then(Instant::now),
    }
}
impl Drop for Span {
    fn drop(&mut self) {
        if let Some(start) = self.start {
            TRACE.with(|trace| {
                if let Some(trace) = trace.borrow_mut().as_mut() {
                    let entry = trace.spans.entry(self.name).or_default();
                    entry.0 += 1;
                    entry.1 += start.elapsed().as_secs_f64() * 1000.0;
                }
            });
        }
    }
}
pub(crate) fn count(name: &'static str, value: u64) {
    TRACE.with(|trace| {
        if let Some(trace) = trace.borrow_mut().as_mut() {
            *trace.counters.entry(name).or_default() += value;
        }
    });
}

pub(crate) fn maximum(name: &'static str, value: u64) {
    TRACE.with(|trace| {
        if let Some(trace) = trace.borrow_mut().as_mut() {
            let entry = trace.counters.entry(name).or_default();
            *entry = (*entry).max(value);
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn disabled_and_nested_commands_do_not_leak() {
        let outer = begin(true);
        count("rows", 3);
        {
            let nested = begin(false);
            count("rows", 8);
            assert!(nested.finish().is_none());
        }
        {
            let _span = span("scan");
        }
        let result = outer.finish().unwrap();
        assert_eq!(result["counters"]["rows"], 3);
        assert_eq!(result["spans"][0]["calls"], 1);
        drop(outer);
        assert!(!enabled());
        assert!(begin(false).finish().is_none());
    }
    #[test]
    fn unwinding_restores_the_previous_trace() {
        let _ = std::panic::catch_unwind(|| {
            let _session = begin(true);
            panic!("test failure");
        });
        assert!(!enabled());
    }
}
