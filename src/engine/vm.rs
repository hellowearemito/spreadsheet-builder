use crate::engine::ast::{
    Cell, CompareOp, Condition, Element, Expr, Expression, ForEachHeader, ForLoop, Format, IfStatement, Modifier, Operator, Row, RowItem
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
                Element::IfStatement(if_statement) => {
                    self.if_statement(if_statement, processor)?;
                }
                Element::ForEachHeader(for_each_header) => {
                    self.for_each_header(for_each_header, processor)?;
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

    pub fn if_statement<'a>(
        &mut self,
        if_statement: &'a IfStatement<'a>,
        processor: &mut impl SheetProcessor,
    ) -> Result<(), SpreadSheetError> {
        let result = self.eval_condition(&if_statement.condition)?;

        if result {
            self.scopes.enter();
            self.run(&if_statement.true_elements, processor)?;
            self.scopes.exit();
        } else {
            self.scopes.enter();
            self.run(&if_statement.false_elements, processor)?;
            self.scopes.exit();
        }
        Ok(())
    }

        pub fn for_each_header(
        &mut self,
        for_each_header: &ForEachHeader,
        processor: &mut impl SheetProcessor,
    ) -> Result<(), SpreadSheetError> {
        let value = self
            .scopes
            .resolve_identifier(for_each_header.variable)
            .cloned()
            .ok_or_else(|| {
                SpreadSheetError::new(format!(
                    "Unresolved identifier: {}",
                    for_each_header.variable
                ))
            })?;

        let Value::Array(arr) = value else {
            return Err(SpreadSheetError::new(format!(
                "header() variable must be an array, got: {}",
                for_each_header.variable
            )));
        };

        let mut cells = Vec::new();

        for item in arr.iter() {
            let Value::Array(tuple) = item else {
                return Err(SpreadSheetError::new(
                    "header() array items must be tuples of [text, span]".to_string(),
                ));
            };

            let text = tuple
                .get(0)
                .ok_or_else(|| SpreadSheetError::new("header tuple missing text field".to_string()))?
                .as_str();

            let span = tuple
                .get(1)
                .ok_or_else(|| SpreadSheetError::new("header tuple missing span field".to_string()))?
                .as_f64() as u16;

            let span = span.max(1);

              cells.push(RowItem::Cell(Cell {
                cell_type: crate::engine::ast::CellType::Str,
                value: Expr::Primary(Expression::Value(Value::String(text))),
                format: for_each_header.format,
                colspan: span,
                rowspan: 1,
                image_mode: None,
            }));
        }

        let row = Row { cells };
        processor.process(&Element::Row(row))?;

        Ok(())
    }

    pub fn eval_condition(&self, condition: &Condition) -> Result<bool, SpreadSheetError> {
        let lhs = self.resolve_expr(&condition.lhs)?;

        match &condition.op {
            None => Ok(lhs.as_bool()),
            Some((op, rhs_expr)) => {
                let rhs = self.resolve_expr(rhs_expr)?;
                let result = match op {
                    CompareOp::Eq  => lhs.eq(&rhs),
                    CompareOp::Neq => !lhs.eq(&rhs),
                    CompareOp::Lt  => lhs.lt(&rhs),
                    CompareOp::Gt  => lhs.gt(&rhs),
                    CompareOp::Lte => lhs.lt(&rhs) || lhs.eq(&rhs),
                    CompareOp::Gte => lhs.gt(&rhs) || lhs.eq(&rhs),
                };
                Ok(result)
            }
        }
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

    pub fn resolve<'b>(&mut self, row: &'b Row) -> Result<Row<'b>, SpreadSheetError> {
        let mut cells = Vec::new();
        for item in &row.cells {
            match item {
                RowItem::Cell(cell) => {
                    let v = self.resolve_expr(&cell.value)?;
                    cells.push(RowItem::Cell(Cell {
                        cell_type: cell.cell_type,
                        value: Expr::Primary(Expression::Value(v)),
                        format: cell.format,
                        colspan: cell.colspan,
                        rowspan: cell.rowspan,
                        image_mode: cell.image_mode,
                    }));
                }
                RowItem::ForEachCell(for_each) => {
                    let value = self.resolve_expression(&for_each.expression)?;
                    if let Value::Array(arr) = value {
                        for (i, v) in arr.iter().enumerate() {
                            self.scopes.enter();
                            self.scopes.top.define("index", Value::Integer(i as i64));
                            self.scopes.top.define(&for_each.variable[1..], v.clone());
                            let resolved_val = self.resolve_expr(&for_each.cell.value)?;
                            cells.push(RowItem::Cell(Cell {
                                cell_type: for_each.cell.cell_type,
                                value: Expr::Primary(Expression::Value(resolved_val)),
                                format: for_each.cell.format,
                                colspan: for_each.cell.colspan,
                                rowspan: for_each.cell.rowspan,
                                image_mode: for_each.cell.image_mode,
                            }));
                            self.scopes.exit();
                        }
                    }
                }
            }
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
