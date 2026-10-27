//! Standalone framed-JSON workbook helper for the independently packaged MCP.
use serde_json::{Value, json};
use std::io::{self, Read, Write};
const REPLY_LIMIT: usize = 16 * 1024 * 1024;
const OVERSIZED_REPLY: &str = "Native workbook reply exceeds transport bound";
fn response_bytes(id: &Value, result: workbook_engine::Result<Value>) -> serde_json::Result<Vec<u8>> {
    let response = match result {
        Ok(result) => json!({"id":id,"ok":true,"result":result}),
        Err(error) => match workbook_engine::split_code(&error) {
            (Some(code), message) => json!({"id":id,"ok":false,"error":message,"code":code}),
            (None, message) => json!({"id":id,"ok":false,"error":message}),
        },
    };
    let bytes = serde_json::to_vec(&response)?;
    if bytes.len() > REPLY_LIMIT {
        // Refusing a large read is recoverable: retain the child and accepted
        // workbook state so its caller can retry a smaller range. Never include
        // workbook content in the bounded error.
        serde_json::to_vec(&json!({"id":id,"ok":false,"error":OVERSIZED_REPLY}))
    } else {
        Ok(bytes)
    }
}
/// Requests up to the size a host's own frames allow (the app core's is 16 MiB).
const REQUEST_LIMIT: usize = 16 * 1024 * 1024;
/// Formula work recurses deeply; run on the stack hosts gave the linked engine.
const STACK: usize = 64 * 1024 * 1024;
fn main() -> Result<(), String> {
    std::thread::Builder::new()
        .name("workbook-engine".into())
        .stack_size(STACK)
        .spawn(|| serve().map_err(|error| error.to_string()))
        .map_err(|error| error.to_string())?
        .join()
        .unwrap_or_else(|_| Err("Workbook helper stopped unexpectedly".into()))
}
fn serve() -> Result<(), Box<dyn std::error::Error>> {
    let mut input = io::stdin().lock();
    let mut output = io::stdout().lock();
    loop {
        let mut header = [0u8; 4];
        match input.read_exact(&mut header) {
            Ok(()) => {}
            Err(e) if e.kind() == io::ErrorKind::UnexpectedEof => break,
            Err(e) => return Err(e.into()),
        }
        let size = u32::from_be_bytes(header) as usize;
        if size == 0 || size > REQUEST_LIMIT {
            return Err("Invalid frame size".into());
        }
        let mut bytes = vec![0; size];
        input.read_exact(&mut bytes)?;
        let request: Value = serde_json::from_slice(&bytes)?;
        let bytes = response_bytes(
            &request["id"],
            workbook_engine::command(request["input"].clone()),
        )?;
        if bytes.len() > REPLY_LIMIT {
            return Err("Reply exceeds transport bound".into());
        }
        output.write_all(&(bytes.len() as u32).to_be_bytes())?;
        output.write_all(&bytes)?;
        output.flush()?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn oversized_reply_keeps_request_id_and_excludes_private_content() {
        let private = "private-workbook-value".repeat(1024 * 1024);
        let encoded =
            response_bytes(&json!(73), Ok(json!({"cells":[["A1",{"value":private}]]}))).unwrap();
        assert!(encoded.len() < 256);
        assert_eq!(
            serde_json::from_slice::<Value>(&encoded).unwrap(),
            json!({"id":73,"ok":false,"error":OVERSIZED_REPLY})
        );
        assert!(
            !String::from_utf8(encoded)
                .unwrap()
                .contains("private-workbook-value")
        );
        assert_eq!(
            serde_json::from_slice::<Value>(
                &response_bytes(&json!(74), Ok(json!({"value":42}))).unwrap()
            )
            .unwrap(),
            json!({"id":74,"ok":true,"result":{"value":42}})
        );
        assert_eq!(
            serde_json::from_slice::<Value>(
                &response_bytes(&json!(75), Err("Unknown workbook version".into())).unwrap()
            )
            .unwrap(),
            json!({"id":75,"ok":false,"error":"Unknown workbook version"})
        );
    }
}
