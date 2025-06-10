use std::fs;
use std::sync::Arc;

fn main() {
    let stream = fs::read_to_string("test.sheet").expect("cannot read file");

    let res = spreadsheet_builder::engine::parser::parse_stream(&stream);

    let dict = spreadsheet_builder::engine::scope::Value::Object(Arc::new(indexmap::indexmap! {
        "inner".into() => spreadsheet_builder::engine::scope::Value::Integer(1234),
    }));

    let arr = spreadsheet_builder::engine::scope::Value::Array(Arc::new(vec![
        spreadsheet_builder::engine::scope::Value::String(String::from("one")),
        spreadsheet_builder::engine::scope::Value::String(String::from("two")),
        spreadsheet_builder::engine::scope::Value::String(String::from("three")),
    ]));

    match res {
        Ok(tree) => {
            let mut vm = spreadsheet_builder::engine::vm::VM::default();
            vm.scopes.top.define(
                "sample",
                spreadsheet_builder::engine::scope::Value::Integer(4765),
            );
            vm.scopes.top.define(
                "red",
                spreadsheet_builder::engine::scope::Value::String("#FF0000".to_string()),
            );
            vm.scopes.top.define("dict", dict);
            vm.scopes.top.define("arr", arr);
            vm.scopes.top.define(
                "bad",
                spreadsheet_builder::engine::scope::Value::String(String::from("sdfsd")),
            );
            let mut writer = spreadsheet_builder::xlsx::XlsxWriter::default();
            let res = vm.run(&tree.elements, &mut writer);
            match res {
                Ok(_) => {
                    // println!("ok");
                }
                Err(e) => {
                    println!("{:?}", e);
                }
            }
            let res = writer.save("output.xlsx");
            match res {
                Ok(_) => {
                    // println!("ok");
                }
                Err(e) => {
                    println!("{:?}", e);
                }
            }
        }
        Err(e) => {
            println!("{:?}", e);
        }
    }
}
