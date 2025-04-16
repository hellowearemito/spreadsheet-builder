pub use crate::engine::ast::*;
pub use crate::engine::scope::Value;
use pest::iterators::Pair;
use pest_derive::Parser;

use crate::engine::diag::SpreadSheetResult;
use pest::pratt_parser::{Assoc, Op, PrattParser};
use pest::Parser;

#[derive(Parser)]
#[grammar = "sheet.pest"]
pub struct SheetParser;

pub fn parse_stream(stream: &str) -> SpreadSheetResult<SyntaxTree> {
    let pairs = SheetParser::parse(Rule::main, stream).map_err(|e| {
        let pos = e.line_col;
        let msg = format!("Syntax error at {:?}, {:?}", pos, e.variant);
        crate::engine::diag::SpreadSheetError::new(msg)
    })?;

    let elements = parse_elements(pairs);

    Ok(SyntaxTree { elements })
}

fn parse_elements(pairs: pest::iterators::Pairs<Rule>) -> Vec<Element> {
    let mut elements = Vec::new();

    for pair in pairs {
        if let Some(element) = parse_element(pair) {
            elements.push(element);
        }
    }

    elements
}

fn parse_element(pair: Pair<Rule>) -> Option<Element> {
    match pair.as_rule() {
        Rule::format_declaration => {
            let format = parse_format(pair.into_inner());
            Some(Element::Format(format))
        }
        Rule::sheet => {
            let sheet = parse_sheet(pair.into_inner());
            Some(Element::Sheet(sheet))
        }
        Rule::anchor => {
            let anchor = parse_anchor(pair.into_inner());
            Some(Element::Anchor(anchor))
        }
        Rule::mover => {
            let mover = parse_mover(pair.into_inner());
            Some(Element::Mover(mover))
        }
        Rule::column => {
            let col = parse_column(pair.into_inner());
            Some(Element::Column(col))
        }
        Rule::rowspec => {
            let rowspec = parse_rowspec(pair.into_inner());
            Some(Element::RowSpec(rowspec))
        }
        Rule::cr => Some(Element::Cr(Cr {})),
        Rule::autofit => Some(Element::Autofit(Autofit {})),
        Rule::row => {
            let row = parse_row(pair.into_inner());
            Some(Element::Row(row))
        }
        Rule::for_loop => {
            let for_loop = parse_for_loop(pair.into_inner());
            Some(Element::ForLoop(for_loop))
        }
        _ => None,
    }
}

fn parse_for_loop(pairs: pest::iterators::Pairs<Rule>) -> ForLoop {
    let mut variable = "";
    let mut expression = Expression::Value(Value::Integer(0));

    let mut elements = Vec::new();
    for pair in pairs {
        match pair.as_rule() {
            Rule::variable_identifier => {
                variable = pair.as_str();
            }
            Rule::expression => {
                expression = parse_expression(pair.into_inner());
            }
            _ => {
                if let Some(element) = parse_element(pair) {
                    elements.push(element);
                }
            }
        }
    }
    ForLoop {
        variable,
        expression,
        elements,
    }
}

fn parse_format(pairs: pest::iterators::Pairs<Rule>) -> Format {
    let mut identifier = "";
    let mut modifiers = Vec::new();
    for pair in pairs {
        match pair.as_rule() {
            Rule::format_identifier => {
                identifier = pair.as_str();
            }
            Rule::format_modifier => {
                let mut statement = "";
                let mut expression = Expr::default();
                for modifier_pair in pair.into_inner() {
                    match modifier_pair.as_rule() {
                        Rule::modifier_statement => {
                            statement = modifier_pair.as_str();
                        }
                        Rule::expr => {
                            let pratt = PrattParser::new()
                                .op(Op::infix(Rule::add, Assoc::Left)
                                    | Op::infix(Rule::sub, Assoc::Left))
                                .op(Op::infix(Rule::mul, Assoc::Left)
                                    | Op::infix(Rule::div, Assoc::Left))
                                .op(Op::prefix(Rule::neg));
                            expression = parse_expr(modifier_pair.into_inner(), &pratt);
                        }
                        _ => {}
                    }
                }
                modifiers.push(Modifier {
                    statement,
                    expression,
                });
            }
            _ => {}
        }
    }
    Format {
        identifier,
        modifiers,
    }
}

fn parse_sheet(pairs: pest::iterators::Pairs<Rule>) -> Sheet {
    let mut name = String::from("");
    for pair in pairs {
        if pair.as_rule() == Rule::sheet_identifier {
            name = decode_string(pair.as_str());
        }
    }
    Sheet { name }
}

fn parse_anchor(pairs: pest::iterators::Pairs<Rule>) -> Anchor {
    let mut identifier = "";
    for pair in pairs {
        if pair.as_rule() == Rule::anchor_identifier {
            identifier = pair.as_str();
        }
    }
    Anchor { identifier }
}

fn decode_string(s: &str) -> String {
    let mut slice = s;

    if s.len() >= 2
        && s.as_bytes().first().copied() == Some(b'"')
        && s.as_bytes().get(s.len() - 1).copied() == Some(b'"')
    {
        slice = &s[1..s.len() - 1];
    }

    let chars = slice.chars();

    let mut quoted = false;

    let mut buffer = String::new();

    for c in chars {
        if quoted {
            match c {
                'n' => buffer.push('\n'),
                'r' => buffer.push('\r'),
                't' => buffer.push('\t'),
                '\\' => buffer.push('\\'),
                '"' => buffer.push('"'),
                _ => buffer.push(c),
            }
            quoted = false;
        } else {
            match c {
                '\\' => {
                    quoted = true;
                }
                _ => buffer.push(c),
            }
        }
    }

    buffer
}

fn parse_value(pair: Pair<Rule>) -> Value {
    let mut value = Value::String(String::from(""));
    match pair.as_rule() {
        Rule::number => {
            if pair.as_str().contains('.') || pair.as_str().contains('e') {
                value = Value::Float(pair.as_str().parse().unwrap_or_default());
            } else {
                value = Value::Integer(pair.as_str().parse().unwrap_or_default());
            }
        }
        Rule::string => {
            value = Value::String(decode_string(pair.as_str()));
        }
        _ => {}
    }

    value
}

fn parse_mover(pairs: pest::iterators::Pairs<Rule>) -> Move {
    let mut anchor = None;
    let mut row = 0;
    let mut col = 0;
    for pair in pairs {
        match pair.as_rule() {
            Rule::anchor_identifier => {
                anchor = Some(pair.as_str());
            }
            Rule::mover_x => {
                row = pair.as_str().parse::<i32>().unwrap_or_default();
            }
            Rule::mover_y => {
                col = pair.as_str().parse::<i16>().unwrap_or_default();
            }
            _ => {}
        }
    }
    Move { anchor, row, col }
}

fn parse_column(pairs: pest::iterators::Pairs<Rule>) -> Column {
    let mut unit = "";
    let mut start = 0;
    let mut end = 0;
    let mut width = 0.0;
    let mut number_idx = 0;
    for pair in pairs {
        match pair.as_rule() {
            Rule::number => {
                if number_idx == 0 {
                    start = pair.as_str().parse::<u16>().unwrap_or_default();
                    number_idx += 1;
                } else if number_idx == 1 {
                    end = pair.as_str().parse::<u16>().unwrap_or_default();
                    number_idx += 1;
                } else if number_idx == 2 {
                    width = pair.as_str().parse::<f64>().unwrap_or_default();
                    number_idx += 1;
                }
            }
            Rule::width_unit => {
                unit = pair.as_str();
            }
            _ => {}
        }
    }
    Column {
        start,
        end,
        unit,
        width,
    }
}

fn parse_rowspec(pairs: pest::iterators::Pairs<Rule>) -> RowSpec {
    let mut unit = "";
    let mut start = 0;
    let mut height = 0.0;
    let mut number_idx = 0;
    for pair in pairs {
        match pair.as_rule() {
            Rule::number => {
                if number_idx == 0 {
                    start = pair.as_str().parse::<u32>().unwrap_or_default();
                    number_idx += 1;
                } else if number_idx == 1 {
                    height = pair.as_str().parse::<f64>().unwrap_or_default();
                    number_idx += 1;
                }
            }
            Rule::width_unit => {
                unit = pair.as_str();
            }
            _ => {}
        }
    }
    RowSpec {
        start,
        unit,
        height,
    }
}

fn parse_cell(pairs: pest::iterators::Pairs<Rule>) -> Cell {
    let mut value = Expr::Primary(Expression::Value(Value::Integer(0)));
    let mut format = None;
    let mut cell_type = CellType::Str;
    let mut colspan = 1;
    let mut rowspan = 1;
    let mut image_mode = None;
    for pair in pairs {
        match pair.as_rule() {
            Rule::cell_type => {
                cell_type = match pair.as_str() {
                    "num" => CellType::Num,
                    "str" => CellType::Str,
                    "date" => CellType::Date,
                    "img" => CellType::Image,
                    _ => CellType::Str,
                };
            }
            Rule::format_identifier => {
                format = Some(pair.as_str());
            }
            Rule::expression => {
                value = Expr::Primary(parse_expression(pair.into_inner()));
            }
            Rule::expr => {
                let pratt = PrattParser::new()
                    .op(Op::infix(Rule::add, Assoc::Left) | Op::infix(Rule::sub, Assoc::Left))
                    .op(Op::infix(Rule::mul, Assoc::Left) | Op::infix(Rule::div, Assoc::Left))
                    .op(Op::prefix(Rule::neg));
                value = parse_expr(pair.into_inner(), &pratt);
                // println!("{:?}", value);
            }
            Rule::image_mode => {
                image_mode = Some(pair.as_str());
            }
            Rule::colspan => {
                let pairs = pair.into_inner();
                for pair in pairs {
                    if pair.as_rule() == Rule::number {
                        colspan = pair.as_str().parse().unwrap_or(1);
                    }
                }
            }
            Rule::rowspan => {
                let pairs = pair.into_inner();
                for pair in pairs {
                    if pair.as_rule() == Rule::number {
                        rowspan = pair.as_str().parse().unwrap_or(1);
                    }
                }
            }
            _ => {}
        }
    }

    Cell {
        cell_type,
        value,
        format,
        colspan,
        rowspan,
        image_mode,
    }
}

fn parse_expression(pairs: pest::iterators::Pairs<Rule>) -> Expression {
    for pair in pairs {
        match pair.as_rule() {
            Rule::number | Rule::string => {
                let value = parse_value(pair);
                return Expression::Value(value);
            }
            Rule::variable_identifier => return Expression::Identifier(pair.as_str()),
            _ => {}
        }
    }

    Expression::Value(Value::Integer(0))
}

fn parse_expr<'a>(pairs: pest::iterators::Pairs<'a, Rule>, pratt: &PrattParser<Rule>) -> Expr<'a> {
    pratt
        .map_primary(|primary| match primary.as_rule() {
            Rule::expression => Expr::Primary(parse_expression(primary.into_inner())),
            Rule::expr => parse_expr(primary.into_inner(), pratt), // from "(" ~ expr ~ ")"
            _ => unreachable!(),
        })
        .map_prefix(|op, rhs| match op.as_rule() {
            Rule::neg => Expr::Prefix(Operator::Neg, Box::new(rhs)),
            _ => unreachable!(),
        })
        .map_infix(|lhs, op, rhs| match op.as_rule() {
            Rule::add => Expr::Infix(Operator::Add, Box::new(lhs), Box::new(rhs)),
            Rule::sub => Expr::Infix(Operator::Sub, Box::new(lhs), Box::new(rhs)),
            Rule::mul => Expr::Infix(Operator::Mul, Box::new(lhs), Box::new(rhs)),
            Rule::div => Expr::Infix(Operator::Div, Box::new(lhs), Box::new(rhs)),
            _ => unreachable!(),
        })
        .parse(pairs)
}

fn parse_row(pairs: pest::iterators::Pairs<Rule>) -> Row {
    let mut cells = Vec::new();
    for pair in pairs {
        if pair.as_rule() == Rule::cell {
            let cell = parse_cell(pair.into_inner());
            cells.push(cell);
        }
    }
    Row { cells }
}
