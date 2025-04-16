use crate::engine::ast::{CellType, Element, Row};
use crate::engine::diag::SpreadSheetError;
use crate::engine::vm::SheetProcessor;
use ecow::EcoString;
use indexmap::IndexMap;
use rust_xlsxwriter::{
    ExcelDateTime, Format, FormatAlign, FormatBorder, FormatScript, FormatUnderline, Image,
    Workbook, Worksheet, XlsxError,
};

pub struct XlsxWriter {
    pub workbook: Workbook,
    pub worksheet: Option<Worksheet>,
    pub row: u32,
    pub col: u16,
    pub anchors: IndexMap<EcoString, (u32, u16)>,
    pub formats: IndexMap<EcoString, Format>,
    pub default_format: Format,
    pub date_format: Format,
    pub number_format: Format,
}

impl Default for XlsxWriter {
    fn default() -> Self {
        XlsxWriter {
            workbook: Workbook::new(),
            worksheet: None,
            row: 0,
            col: 0,
            anchors: IndexMap::new(),
            formats: IndexMap::new(),
            default_format: Format::new(),
            date_format: Format::new().set_num_format("dd/mm/yyyy hh:mm"),
            number_format: Format::new().set_num_format("0.00"),
        }
    }
}

impl XlsxWriter {
    pub fn save(&mut self, path: &str) -> Result<(), XlsxError> {
        if let Some(sheet) = self.worksheet.take() {
            // println!("pushing worksheet: {:?}", sheet.name());
            self.workbook.push_worksheet(sheet);
        }
        self.workbook.save(path)
    }

    pub fn process_internal(&mut self, item: &Element) -> Result<(), XlsxError> {
        // println!("processing item {:?}", item);
        match item {
            Element::Sheet(sheet) => {
                let sheet_name = &sheet.name;
                if let Some(sheet) = self.worksheet.take() {
                    self.workbook.push_worksheet(sheet);
                }
                let mut sheet = Worksheet::new();
                sheet.set_name(sheet_name)?;
                self.worksheet = Some(sheet);
                self.row = 0;
                self.col = 0;
            }
            Element::Row(row) => {
                self.process_row(row)?;
            }
            Element::Anchor(anchor) => {
                let (row, col) = (self.row, self.col);
                self.anchors
                    .insert(EcoString::from(anchor.identifier), (row, col));
            }
            Element::Format(format) => {
                self.process_format(format)?;
            }
            Element::Mover(mover) => {
                if let Some(anchor) = mover.anchor {
                    if let Some((a_row, a_col)) = self.anchors.get(anchor) {
                        self.row = a_row.checked_add_signed(mover.row).unwrap_or_default();
                        self.col = a_col.checked_add_signed(mover.col).unwrap_or_default();
                    }
                } else {
                    self.row = self.row.checked_add_signed(mover.row).unwrap_or_default();
                    self.col = self.col.checked_add_signed(mover.col).unwrap_or_default();
                }
            }
            Element::Cr(_) => {
                self.row += 1;
                self.col = 0;
            }
            Element::Autofit(_) => {
                if let Some(sheet) = self.worksheet.as_mut() {
                    sheet.autofit();
                }
            }
            Element::Column(column) => {
                if let Some(sheet) = self.worksheet.as_mut() {
                    if column.unit == "chars" {
                        sheet.set_column_range_width(column.start, column.end, column.width)?;
                    } else {
                        sheet.set_column_range_width_pixels(
                            column.start,
                            column.end,
                            column.width as u16,
                        )?;
                    }
                }
            }
            Element::RowSpec(rowspec) => {
                if let Some(sheet) = self.worksheet.as_mut() {
                    if rowspec.unit == "chars" {
                        sheet.set_row_height(rowspec.start, rowspec.height)?;
                    } else {
                        sheet.set_row_height_pixels(rowspec.start, rowspec.height as u16)?;
                    }
                }
            }
            _ => {}
        }

        Ok(())
    }

    pub fn process_row(&mut self, row: &Row) -> Result<(), XlsxError> {
        if self.worksheet.is_some() {
            let sheet = self.worksheet.as_mut().unwrap();
            let save_col = self.col;
            for cell in &row.cells {
                let format = if let Some(f) = cell.format {
                    if let Some(f) = self.formats.get(f) {
                        f
                    } else {
                        match cell.cell_type {
                            CellType::Num => &self.number_format,
                            CellType::Date => &self.date_format,
                            _ => &self.default_format,
                        }
                    }
                } else {
                    match cell.cell_type {
                        CellType::Num => &self.number_format,
                        CellType::Date => &self.date_format,
                        _ => &self.default_format,
                    }
                };

                if cell.colspan > 1 || cell.rowspan > 1 {
                    let end_row = self.row + cell.rowspan as u32 - 1;
                    let end_col = self.col + cell.colspan - 1;
                    match cell.cell_type {
                        CellType::Str => {
                            sheet.merge_range(
                                self.row,
                                self.col,
                                end_row,
                                end_col,
                                &cell.value.as_str(),
                                format,
                            )?;
                        }
                        _ => {
                            sheet.merge_range(self.row, self.col, end_row, end_col, "", format)?;
                        }
                    }
                }

                match cell.cell_type {
                    CellType::Num => {
                        sheet.write_number_with_format(
                            self.row,
                            self.col,
                            cell.value.as_f64(),
                            format,
                        )?;
                    }
                    CellType::Str => {
                        if cell.colspan == 1 && cell.rowspan == 1 {
                            sheet.write_string_with_format(
                                self.row,
                                self.col,
                                cell.value.as_str(),
                                format,
                            )?;
                        }
                    }
                    CellType::Date => {
                        sheet.write_with_format(
                            self.row,
                            self.col,
                            ExcelDateTime::parse_from_str(&cell.value.as_str())?,
                            format,
                        )?;
                    }
                    CellType::Image => {
                        let image_mode = cell.image_mode.unwrap_or("embed");
                        let image = Image::new(cell.value.as_str())?;
                        match image_mode {
                            "embed" => {
                                sheet.insert_image_fit_to_cell(self.row, self.col, &image, true)?;
                            }
                            "insert" => {
                                sheet.insert_image(self.row, self.col, &image)?;
                            }
                            _ => {}
                        }
                    }
                }
                self.col += cell.colspan;
            }

            self.row += 1;
            self.col = save_col;
        }

        Ok(())
    }

    pub fn process_format(&mut self, format: &crate::engine::ast::Format) -> Result<(), XlsxError> {
        let mut f = Format::new();
        for modifier in &format.modifiers {
            let param_string = modifier.expression.as_str();
            let param = param_string.as_str();

            match modifier.statement {
                "bold" => {
                    f = f.set_bold();
                }
                "italic" => {
                    f = f.set_italic();
                }
                "underline" => {
                    f = f.set_underline(FormatUnderline::Single);
                }
                "strikethrough" => {
                    f = f.set_font_strikethrough();
                }
                "super" => {
                    f = f.set_font_script(FormatScript::Superscript);
                }
                "sub" => {
                    f = f.set_font_script(FormatScript::Subscript);
                }
                "color" => {
                    f = f.set_font_color(param);
                }
                "num" => {
                    f = f.set_num_format(param);
                }
                "align" => {
                    match param {
                        "left" => f = f.set_align(FormatAlign::Left),
                        "right" => f = f.set_align(FormatAlign::Right),
                        "center" => f = f.set_align(FormatAlign::Center),
                        "top" => f = f.set_align(FormatAlign::Top),
                        "bottom" => f = f.set_align(FormatAlign::Bottom),
                        "verticalcenter" => f = f.set_align(FormatAlign::VerticalCenter),
                        _ => {}
                    };
                }
                "indent" => {
                    if let Ok(indent) = param.parse::<u8>() {
                        f = f.set_indent(indent);
                    }
                }
                "font_name" => f = f.set_font_name(param),
                "font_size" => {
                    if let Ok(size) = param.parse::<f64>() {
                        f = f.set_font_size(size);
                    }
                }
                "background_color" => f = f.set_background_color(param),
                "border" => {
                    let border = interpret_border(param);
                    f = f.set_border(border);
                }
                "border_top" => {
                    let border = interpret_border(param);
                    f = f.set_border_top(border);
                }
                "border_bottom" => {
                    let border = interpret_border(param);
                    f = f.set_border_bottom(border);
                }
                "border_left" => {
                    let border = interpret_border(param);
                    f = f.set_border_left(border);
                }
                "border_right" => {
                    let border = interpret_border(param);
                    f = f.set_border_right(border);
                }
                "border_color" => f = f.set_border_color(param),
                "border_top_color" => f = f.set_border_top_color(param),
                "border_bottom_color" => f = f.set_border_bottom_color(param),
                "border_left_color" => f = f.set_border_left_color(param),
                "border_right_color" => f = f.set_border_right_color(param),
                _ => {}
            }
        }
        self.formats.insert(EcoString::from(format.identifier), f);
        Ok(())
    }
}

impl SheetProcessor for XlsxWriter {
    fn process(&mut self, item: &Element) -> Result<(), SpreadSheetError> {
        self.process_internal(item).map_err(handle_error)
    }
}

fn handle_error(e: XlsxError) -> SpreadSheetError {
    let msg = format!("{:?}", e);
    SpreadSheetError::new(msg)
}

fn interpret_border(border: &str) -> FormatBorder {
    match border {
        "none" => rust_xlsxwriter::FormatBorder::None,
        "thin" => rust_xlsxwriter::FormatBorder::Thin,
        "medium" => rust_xlsxwriter::FormatBorder::Medium,
        "dashed" => rust_xlsxwriter::FormatBorder::Dashed,
        "dotted" => rust_xlsxwriter::FormatBorder::Dotted,
        "thick" => rust_xlsxwriter::FormatBorder::Thick,
        "double" => rust_xlsxwriter::FormatBorder::Double,
        "hair" => rust_xlsxwriter::FormatBorder::Hair,
        "medium_dashed" => rust_xlsxwriter::FormatBorder::MediumDashed,
        "dash_dot" => rust_xlsxwriter::FormatBorder::DashDot,
        "medium_dash_dot" => rust_xlsxwriter::FormatBorder::MediumDashDot,
        "dash_dot_dot" => rust_xlsxwriter::FormatBorder::DashDotDot,
        "medium_dash_dot_dot" => rust_xlsxwriter::FormatBorder::MediumDashDotDot,
        "slant_dash_dot" => rust_xlsxwriter::FormatBorder::SlantDashDot,
        _ => rust_xlsxwriter::FormatBorder::Thin,
    }
}
