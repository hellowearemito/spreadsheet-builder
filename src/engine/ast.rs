use crate::engine::scope::Value;

#[derive(Debug)]
pub struct Modifier<'a> {
    pub statement: &'a str,
    pub expression: Expr<'a>,
}

#[derive(Debug)]
pub struct Format<'a> {
    pub identifier: &'a str,
    pub modifiers: Vec<Modifier<'a>>,
}

#[derive(Debug)]
pub struct Anchor<'a> {
    pub identifier: &'a str,
}

#[derive(Debug)]
pub struct Sheet {
    pub name: String,
}

#[derive(Debug)]
pub struct Move<'a> {
    pub anchor: Option<&'a str>,
    pub row: i32,
    pub col: i16,
}

#[derive(Debug)]
pub struct ForEachCell<'a> {
    pub variable: &'a str,
    pub expression: Expression<'a>,
    pub cell: Cell<'a>,
}

#[derive(Debug)]
pub enum RowItem<'a> {
    Cell(Cell<'a>),
    ForEachCell(ForEachCell<'a>),
}

#[derive(Debug)]
pub struct Row<'a> {
    pub cells: Vec<RowItem<'a>>,
}

#[derive(Debug, Copy, Clone)]
pub enum CellType {
    Num,
    Str,
    Date,
    Image,
    Bool,
}

#[derive(Debug)]
pub struct Cell<'a> {
    pub cell_type: CellType,
    pub value: Expr<'a>,
    pub format: Option<&'a str>,
    pub colspan: u16,
    pub rowspan: u16,
    pub image_mode: Option<&'a str>,
}

#[derive(Debug)]
pub struct Cr {}

#[derive(Debug)]
pub struct Autofit {}

#[derive(Debug)]
pub enum Element<'a> {
    Format(Format<'a>),
    Sheet(Sheet),
    Anchor(Anchor<'a>),
    Row(Row<'a>),
    Mover(Move<'a>),
    ForLoop(ForLoop<'a>),
    IfStatement(IfStatement<'a>),
    Cr(Cr),
    Autofit(Autofit),
    Column(Column<'a>),
    RowSpec(RowSpec<'a>),
}

#[derive(Debug)]
pub struct SyntaxTree<'a> {
    pub elements: Vec<Element<'a>>,
}

#[derive(Debug)]
pub enum Expression<'a> {
    Value(Value),
    Identifier(&'a str),
}

#[derive(Debug)]
pub enum Operator {
    Add,
    Sub,
    Mul,
    Div,
    Neg,
}

#[derive(Debug)]
pub enum Expr<'a> {
    Primary(Expression<'a>),
    Infix(Operator, Box<Expr<'a>>, Box<Expr<'a>>),
    Prefix(Operator, Box<Expr<'a>>),
}

impl Default for Expr<'_> {
    fn default() -> Self {
        Expr::Primary(Expression::Value(Value::String("".to_string())))
    }
}

#[derive(Debug)]
pub struct ForLoop<'a> {
    pub variable: &'a str,
    pub expression: Expression<'a>,
    pub elements: Vec<Element<'a>>,
}

#[derive(Debug)]
pub enum CompareOp {
    Eq,
    Neq,
    Lt,
    Gt,
    Lte,
    Gte,
}

#[derive(Debug)]
pub struct Condition<'a> {
    pub lhs: Expr<'a>,
    pub op: Option<(CompareOp, Expr<'a>)>,
}

#[derive(Debug)]
pub struct IfStatement<'a> {
    pub condition: Condition<'a>,
    pub true_elements: Vec<Element<'a>>,
    pub false_elements: Vec<Element<'a>>,
}

#[derive(Debug)]
pub struct Column<'a> {
    pub start: u16,
    pub end: u16,
    pub unit: &'a str,
    pub width: f64,
}

#[derive(Debug)]
pub struct RowSpec<'a> {
    pub start: u32,
    pub unit: &'a str,
    pub height: f64,
}

impl Expression<'_> {
    pub fn as_str(&self) -> String {
        match self {
            Expression::Value(v) => v.as_str(),
            _ => String::from(""),
        }
    }

    pub fn as_f64(&self) -> f64 {
        match self {
            Expression::Value(v) => v.as_f64(),
            _ => 0.0,
        }
    }

    pub fn as_bool(&self) -> bool {
        match self {
            Expression::Value(v) => v.as_bool(),
            _ => false,
        }
    }
}

impl Expr<'_> {
    pub fn as_str(&self) -> String {
        match self {
            Expr::Primary(v) => v.as_str(),
            _ => String::from(""),
        }
    }

    pub fn as_f64(&self) -> f64 {
        match self {
            Expr::Primary(v) => v.as_f64(),
            _ => 0.0,
        }
    }

    pub fn as_bool(&self) -> bool {
        match self {
            Expr::Primary(v) => v.as_bool(),
            _ => false,
        }
    }
}
