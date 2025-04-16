use crate::engine::ast::{
    Cell, Element, Expr, Expression, ForLoop, Format, Modifier, Operator, Row,
};
use crate::engine::diag::SpreadSheetError;
use crate::engine::scope::{Scopes, Value};

pub trait SheetProcessor {
    fn process(&mut self, item: &Element) -> Result<(), SpreadSheetError>;
}

pub struct VM {
    pub scopes: Scopes,
}

impl Default for VM {
    fn default() -> Self {
        Self {
            scopes: Scopes::new(),
        }
    }
}

impl VM {
    pub fn run<'a>(
        &mut self,
        items: &'a [Element<'a>],
        processor: &mut impl SheetProcessor,
    ) -> Result<(), SpreadSheetError> {
        for item in items {
            match item {
                Element::Format(format) => {
                    let format = self.resolve_format(format)?;
                    processor.process(&Element::Format(format))?
                }
                Element::Row(row) => {
                    let row = self.resolve(row)?;
                    processor.process(&Element::Row(row))?;
                }
                Element::ForLoop(for_loop) => {
                    self.for_loop(for_loop, processor)?;
                }
                _ => {
                    processor.process(item)?;
                }
            }
        }
        Ok(())
    }

    pub fn for_loop<'a>(
        &mut self,
        for_loop: &'a ForLoop<'a>,
        processor: &mut impl SheetProcessor,
    ) -> Result<(), SpreadSheetError> {
        let value = self.resolve_expression(&for_loop.expression)?;
        if let Value::Array(arr) = value {
            for (i, v) in arr.iter().enumerate() {
                self.scopes.enter();
                self.scopes.top.define("index", Value::Integer(i as i64));
                self.scopes.top.define(&for_loop.variable[1..], v.clone());
                self.run(&for_loop.elements, processor)?;
                self.scopes.exit();
            }
        }
        Ok(())
    }

    pub fn resolve_expression(&self, expression: &Expression) -> Result<Value, SpreadSheetError> {
        let v = match expression {
            Expression::Value(v) => v.clone(),
            Expression::Identifier(id) => {
                if let Some(v) = self.scopes.resolve(Expression::Identifier(id)) {
                    v
                } else {
                    return Err(SpreadSheetError::new(format!(
                        "Unresolved identifier: {}",
                        id
                    )));
                }
            }
        };

        Ok(v)
    }

    pub fn handle(op: &Operator, lhs: &Value, rhs: &Value) -> Result<Value, SpreadSheetError> {
        match op {
            Operator::Add => lhs.add(rhs),
            Operator::Sub => lhs.sub(rhs),
            Operator::Mul => lhs.mul(rhs),
            Operator::Div => lhs.div(rhs),
            _ => Err(SpreadSheetError::new("Invalid infix operator".to_string())),
        }
    }

    pub fn resolve_expr(&self, expr: &Expr) -> Result<Value, SpreadSheetError> {
        match expr {
            Expr::Infix(op, lhs, rhs) => {
                if let Expr::Primary(Expression::Identifier(id)) = lhs.as_ref() {
                    if let Some(lhs_v) = self.scopes.resolve_identifier(id) {
                        if let Expr::Primary(Expression::Identifier(id2)) = rhs.as_ref() {
                            if let Some(rhs_v) = self.scopes.resolve_identifier(id2) {
                                return Self::handle(op, lhs_v, rhs_v);
                            }
                        }

                        return Self::handle(op, lhs_v, &self.resolve_expr(rhs.as_ref())?);
                    }
                }

                let lhs = self.resolve_expr(lhs.as_ref())?;

                if let Expr::Primary(Expression::Identifier(id2)) = rhs.as_ref() {
                    if let Some(rhs_v) = self.scopes.resolve_identifier(id2) {
                        return Self::handle(op, &lhs, rhs_v);
                    }
                }

                Self::handle(op, &lhs, &self.resolve_expr(rhs.as_ref())?)
            }
            Expr::Prefix(op, expr) => {
                let expr = self.resolve_expr(expr.as_ref())?;
                match op {
                    Operator::Neg => expr.neg(),
                    _ => Err(SpreadSheetError::new("Invalid prefix operator".to_string())),
                }
            }
            Expr::Primary(expr) => self.resolve_expression(expr),
        }
    }

    pub fn resolve<'b>(&self, row: &'b Row) -> Result<Row<'b>, SpreadSheetError> {
        let mut cells = Vec::new();
        for cell in &row.cells {
            let v = self.resolve_expr(&cell.value)?;
            cells.push(Cell {
                cell_type: cell.cell_type,
                value: Expr::Primary(Expression::Value(v)),
                format: cell.format,
                colspan: cell.colspan,
                rowspan: cell.rowspan,
                image_mode: cell.image_mode,
            });
        }
        Ok(Row { cells })
    }

    pub fn resolve_format<'b>(&self, format: &'b Format) -> Result<Format<'b>, SpreadSheetError> {
        let mut modifiers = Vec::new();
        for modifier in &format.modifiers {
            let v = self.resolve_expr(&modifier.expression)?;
            modifiers.push(Modifier {
                statement: modifier.statement,
                expression: Expr::Primary(Expression::Value(v)),
            });
        }
        Ok(Format {
            identifier: format.identifier,
            modifiers,
        })
    }
}
