use std::collections::BTreeMap;

const MAX_EVENTS: usize = 256;
const MAX_BYTES: usize = 32 * 1024 * 1024;

/// Track events until the trusted renderer finishes dispatching their callbacks.
/// A stopped renderer cannot create an unbounded native/WebKit event backlog.
#[derive(Default)]
pub struct EventBudget {
    next: u64,
    pending: BTreeMap<u64, usize>,
    bytes: usize,
}
impl EventBudget {
    pub fn reserve(&mut self, bytes: usize) -> Result<u64, String> {
        if self.pending.len() >= MAX_EVENTS || bytes > MAX_BYTES.saturating_sub(self.bytes) {
            return Err(
                "unavailable: Renderer event budget exceeded; reopen the app to resynchronize"
                    .into(),
            );
        }
        self.next += 1;
        self.pending.insert(self.next, bytes);
        self.bytes += bytes;
        Ok(self.next)
    }
    pub fn acknowledge(&mut self, through: u64) -> Result<(), String> {
        if through > self.next {
            return Err("validation: Unknown event receipt".into());
        }
        while let Some((&sequence, &bytes)) = self.pending.first_key_value() {
            if sequence > through {
                break;
            }
            self.bytes -= bytes;
            self.pending.remove(&sequence);
        }
        Ok(())
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn stalled_renderer_is_bounded_and_receipts_release_only_delivered_events() {
        let mut budget = EventBudget::default();
        for expected in 1..=MAX_EVENTS {
            assert_eq!(budget.reserve(100).unwrap(), expected as u64);
        }
        assert!(budget.reserve(1).is_err());
        assert!(budget.acknowledge(1000).is_err());
        budget.acknowledge(128).unwrap();
        assert_eq!(budget.pending.len(), 128);
        assert_eq!(budget.bytes, 12800);
        budget.acknowledge(128).unwrap(); // Duplicate receipt is harmless.
        assert_eq!(budget.reserve(100).unwrap(), 257);
        budget.acknowledge(257).unwrap();
        assert_eq!(budget.bytes, 0);
    }
    #[test]
    fn bytes_are_bounded_even_with_few_large_events() {
        let mut budget = EventBudget::default();
        assert!(budget.reserve(MAX_BYTES + 1).is_err());
        let id = budget.reserve(MAX_BYTES).unwrap();
        assert!(budget.reserve(1).is_err());
        budget.acknowledge(id).unwrap();
        assert!(budget.reserve(MAX_BYTES).is_ok());
    }
}
