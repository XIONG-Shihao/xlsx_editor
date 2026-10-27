use std::io::Read;
fn main() {
    let mut input = String::new();
    std::io::stdin().read_to_string(&mut input).unwrap();
    let result = workbook_engine::command(serde_json::from_str(&input).unwrap());
    println!(
        "{}",
        match result {
            Ok(value) => serde_json::json!({"ok":true,"result":value}),
            Err(error) => serde_json::json!({"ok":false,"error":error}),
        }
    );
}
