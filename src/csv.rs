use crate::engine::ast::{CellType, Element, Row, RowItem};
use crate::engine::diag::SpreadSheetError;
use crate::engine::vm::SheetProcessor;
use csv::{Writer, WriterBuilder};
use std::fs::File;
use std::path::Path;

pub struct CsvWriter {
    pub writer: Writer<File>,
}

impl CsvWriter {
    pub fn from_path<P: AsRef<Path>>(path: P, delimiter: u8) -> Result<Self, csv::Error> {
        let writer = WriterBuilder::new().delimiter(delimiter).from_path(path)?;
        Ok(CsvWriter { writer })
    }

    pub fn save(&mut self) -> Result<(), csv::Error> {
        self.writer.flush()?;
        Ok(())
    }

    pub fn process_internal(&mut self, item: &Element) -> Result<(), csv::Error> {
        // println!("processing item {:?}", item);
        if let Element::Row(row) = item {
            self.process_row(row)?;
        }

        Ok(())
    }

    pub fn process_row(&mut self, row: &Row) -> Result<(), csv::Error> {
        for item in &row.cells {
            let cell = match item {
                RowItem::Cell(cell) => cell,
                RowItem::ForEachCell(_) => continue,
            };
            match cell.cell_type {
                CellType::Num => {
                    self.writer.write_field(cell.value.as_f64().to_string())?;
                }
                CellType::Str => {
                    self.writer.write_field(cell.value.as_str())?;
                }
                CellType::Date => {
                    self.writer.write_field(cell.value.as_str())?;
                }
                CellType::Image => {
                    // ignore
                }
                CellType::Bool => {
                    self.writer.write_field(cell.value.as_str())?;
                }
            }
            for _ in 1..cell.colspan {
                self.writer.write_field("")?;
            }
        }

        self.writer.write_record(None::<&[u8]>)?;

        Ok(())
    }
}

impl SheetProcessor for CsvWriter {
    fn process(&mut self, item: &Element) -> Result<(), SpreadSheetError> {
        self.process_internal(item).map_err(handle_error)
    }
}

fn handle_error(e: csv::Error) -> SpreadSheetError {
    let msg = format!("{:?}", e);
    SpreadSheetError::new(msg)
}
