use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::{env, io::Write};

fn main() -> std::io::Result<()> {
    let file_path = env::args().nth(1).expect("Please provide an excel");

    let sheets = create(file_path);
    let sql_string = to_sql(sheets);
    let sql: String = sql_string.join(" ");
    let mut file = std::fs::File::create("data.sql")?;

    file.write_all(sql.as_bytes())?;
    Ok(())
}

#[derive(Debug)]
struct Sheet {
    name: String,
    headers: Vec<String>,
    data: Vec<MappedData>,
}

#[derive(Debug)]
struct MappedData {
    key: String,
    value: String,
}

impl Sheet {
    fn new(item: &(String, Range<DataType>)) -> Self {
        let headers: Vec<String> = Self::create_headers(&item.1);
        let mapped_data: Vec<MappedData> = Self::map_data(&item.1, &headers);
        Self {
            name: item.0.to_owned(),
            headers,
            data: mapped_data,
        }
    }

    fn create_headers(data: &Range<DataType>) -> Vec<String> {
        let mut ret: Vec<String> = vec![];

        for i in data.rows().nth(0).unwrap() {
            if let Some(s) = i.get_string() {
                ret.push(s.to_string());
            }
        }
        ret
    }

    fn map_data(data: &Range<DataType>, headers: &Vec<String>) -> Vec<MappedData> {
        let mut ret: Vec<MappedData> = vec![];

        for (index, row) in data.rows().enumerate() {
            if index == 0 {
                continue;
            }
            for i in 0..headers.len() {
                let d = &row[i];
                let mp = MappedData {
                    key: headers[i].to_string(),
                    value: d.to_string(),
                };
                ret.push(mp);
            }
        }
        ret
    }
}

fn create(path: String) -> Vec<Sheet> {
    let mut workbook = open_workbook_auto(&path).unwrap();

    let mut sheets: Vec<Sheet> = vec![];

    for i in workbook.worksheets() {
        let temp = Sheet::new(&i);
        sheets.push(temp);
    }
    sheets
}

fn to_sql(sheets: Vec<Sheet>) -> Vec<String> {
    let mut sql_list = vec![];

    for sheet in sheets.into_iter() {
        let mut builder = sql_builder::SqlBuilder::insert_into(sheet.name);
        for header in sheet.headers.into_iter() {
            builder.field(header);
        }

        let mut str_slice: Vec<String> = vec![];
        for data in sheet.data.into_iter() {
            str_slice.push(data.value);
        }
        builder.values(&str_slice[..]);

        let sql_string = builder.sql().unwrap();
        sql_list.push(sql_string);
    }
    return sql_list;
}
