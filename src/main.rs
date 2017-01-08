extern crate calamine;

use std::env;
use calamine::{Excel, DataType, CellErrorType};

fn help() {
    println!("Usage: git-xlsx-text-conv filename.xlsx");
}

fn textconv(excel_file_path: &String) {
    let mut excel:Excel = Excel::open(excel_file_path).expect("The provided file was not an readable Excel file");
    
    for sheet_name in excel.sheet_names().unwrap() {
        let sheet = excel.worksheet_range(&sheet_name).unwrap();
        for row in sheet.rows() {
            let mut cells: Vec<String> = Vec::new();
            
            for cell in row {
                let str: String = datatype_to_str(&cell)
                    .replace("\\", "\\\\")
                    .replace("\n", "\\n")
                    .replace("\r", "\\r")
                    .replace("\t", "\\t");
                
                cells.push(str);
            }
            
            println!("[{}] {}", sheet_name, cells.join("\t"));
        }
    }
}

fn datatype_to_str(data_type: &DataType) -> String {
    return match *data_type {
        DataType::Bool(true) => "TRUE".to_string(),
        DataType::Bool(false) => "FALSE".to_string(),
        DataType::Int(value) => value.to_string(),
        DataType::Float(value) => value.to_string(),
        DataType::String(ref value) => value.clone(),
        DataType::Error(ref error) => match *error {
            CellErrorType::Div0 => "#DIV/0!".to_string(),
            CellErrorType::NA => "#N/A".to_string(),
            CellErrorType::Name => "#NAME?".to_string(),
            CellErrorType::Null => "#NULL!".to_string(),
            CellErrorType::Num => "#NUM!".to_string(),
            CellErrorType::Ref =>"#REF!".to_string(),
            CellErrorType::Value => "#VALUE!".to_string(),
            CellErrorType::GettingData => "#GETTING_DATA".to_string(),
        }, 
        DataType::Empty => "".to_string(),
    };
}

fn main() {
    let args: Vec<String> = env::args().collect();
    match args.len() {
        2 => textconv(&args[1]),
        _ => help(),
    }
}
