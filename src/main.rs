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
            vm.scopes.top.define(
                "mytrueboolean",
                spreadsheet_builder::engine::scope::Value::Boolean(true),
            );
            vm.scopes.top.define(
                "myfalseboolean",
                spreadsheet_builder::engine::scope::Value::Boolean(false),
            );
            vm.scopes.top.define(
                "myinteger",
                spreadsheet_builder::engine::scope::Value::Integer(3),
            );
            vm.scopes.top.define(
                "mystring",
                spreadsheet_builder::engine::scope::Value::String(String::from("hello")),
            );
            vm.scopes.top.define(
                "zero",
                spreadsheet_builder::engine::scope::Value::Integer(0),
            );
            vm.scopes.top.define(
                "negative",
                spreadsheet_builder::engine::scope::Value::Integer(-5),
            );
            vm.scopes.top.define(
                "largefloat",
                spreadsheet_builder::engine::scope::Value::Float(9.99),
            );
            vm.scopes.top.define(
                "smallfloat",
                spreadsheet_builder::engine::scope::Value::Float(0.01),
            );
            vm.scopes.top.define(
                "emptystring",
                spreadsheet_builder::engine::scope::Value::String(String::from("")),
            );
            vm.scopes.top.define(
                "truestring",
                spreadsheet_builder::engine::scope::Value::String(String::from("true")),
            );
            vm.scopes.top.define(
                "falsestring",
                spreadsheet_builder::engine::scope::Value::String(String::from("false")),
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
