use crate::engine::ast::Expression;
use crate::engine::diag::{SpreadSheetError, SpreadSheetResult};
use ecow::EcoString;
use indexmap::IndexMap;
use serde::de::value::{MapAccessDeserializer, SeqAccessDeserializer};
use serde::de::{Error, MapAccess, SeqAccess, Visitor};
use serde::{Deserialize, Deserializer};
use std::fmt;
use std::fmt::{Debug, Formatter};
use std::sync::Arc;

#[derive(Debug, Clone)]
pub enum Value {
    String(String),
    Integer(i64),
    Float(f64),
    Boolean(bool),
    Array(Arc<Vec<Value>>),
    Object(Arc<IndexMap<EcoString, Value>>),
}

impl Value {
    pub fn as_f64(&self) -> f64 {
        match self {
            Value::Float(f) => *f,
            Value::Integer(i) => *i as f64,
            Value::String(s) => s.parse().unwrap_or(0.0),
            _ => 0.0,
        }
    }

    pub fn as_str(&self) -> String {
        match self {
            Value::String(s) => String::from(s),
            Value::Float(f) => f.to_string(),
            Value::Integer(i) => i.to_string(),
            _ => String::from(""),
        }
    }

    pub fn resolve(&self, path: &mut PathSplitter) -> Option<&Value> {
        if let Some(name) = path.next() {
            match self {
                Value::Object(map) => map.get(name).and_then(|v| v.resolve(path)),
                Value::Array(arr) => {
                    if let Ok(index) = name.parse::<usize>() {
                        arr.get(index).and_then(|v| v.resolve(path))
                    } else {
                        None
                    }
                }
                _ => None,
            }
        } else {
            Some(self)
        }
    }

    pub fn add(&self, rhs: &Value) -> SpreadSheetResult<Value> {
        match self {
            Value::Integer(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Integer(lhs + *rhs)),
                Value::Float(rhs) => Ok(Value::Float(*lhs as f64 + *rhs)),
                Value::String(rhs) => Ok(Value::String(lhs.to_string() + rhs)),
                Value::Array(rhs) => {
                    let n = rhs.len();
                    Ok(Value::Integer(*lhs + n as i64))
                }
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer + object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer + boolean".to_string(),
                )),
            },
            Value::Float(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Float(lhs + *rhs as f64)),
                Value::Float(rhs) => Ok(Value::Float(lhs + *rhs)),
                Value::String(rhs) => Ok(Value::String(lhs.to_string() + rhs)),
                Value::Array(rhs) => {
                    let n = rhs.len();
                    Ok(Value::Float(*lhs + n as f64))
                }
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: float + object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: float + boolean".to_string(),
                )),
            },
            Value::String(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::String(lhs.to_string() + &rhs.to_string())),
                Value::Float(rhs) => Ok(Value::String(lhs.to_string() + &rhs.to_string())),
                Value::String(rhs) => Ok(Value::String(lhs.to_string() + rhs)),
                Value::Array(_) => Err(SpreadSheetError::new(
                    "invalid operation: string + array".to_string(),
                )),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: string + object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: string + boolean".to_string(),
                )),
            },
            Value::Array(lhs) => match rhs {
                Value::Integer(rhs) => {
                    let n = lhs.len();
                    Ok(Value::Integer(n as i64 + rhs))
                }
                Value::Float(rhs) => {
                    let n = lhs.len();
                    Ok(Value::Float(n as f64 + rhs))
                }
                Value::Array(rhs) => Ok(Value::Array(Arc::new(
                    lhs.iter().chain(rhs.iter()).cloned().collect(),
                ))),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: array + object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: array + boolean".to_string(),
                )),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: array + string".to_string(),
                )),
            },
            Value::Boolean(_) => Err(SpreadSheetError::new(
                "invalid operation: boolean + _".to_string(),
            )),
            Value::Object(_) => Err(SpreadSheetError::new(
                "invalid operation: object + _".to_string(),
            )),
        }
    }

    pub fn sub(&self, rhs: &Value) -> SpreadSheetResult<Value> {
        match self {
            Value::Integer(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Integer(lhs - rhs)),
                Value::Float(rhs) => Ok(Value::Float(*lhs as f64 - rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer - string".to_string(),
                )),
                Value::Array(rhs) => {
                    let n = rhs.len();
                    Ok(Value::Integer(*lhs - n as i64))
                }
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer - object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer - boolean".to_string(),
                )),
            },
            Value::Float(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Float(lhs - *rhs as f64)),
                Value::Float(rhs) => Ok(Value::Float(lhs - rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: float - string".to_string(),
                )),
                Value::Array(rhs) => {
                    let n = rhs.len();
                    Ok(Value::Float(*lhs - n as f64))
                }
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: float - object".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: float - boolean".to_string(),
                )),
            },
            Value::String(_) => Err(SpreadSheetError::new(
                "invalid operation: string - _".to_string(),
            )),
            Value::Boolean(_) => Err(SpreadSheetError::new(
                "invalid operation: boolean - _".to_string(),
            )),
            Value::Object(_) => Err(SpreadSheetError::new(
                "invalid operation: object - _".to_string(),
            )),
            Value::Array(_) => Err(SpreadSheetError::new(
                "invalid operation: array - _".to_string(),
            )),
        }
    }

    pub fn neg(&self) -> SpreadSheetResult<Value> {
        match self {
            Value::Integer(i) => Ok(Value::Integer(-i)),
            Value::Float(f) => Ok(Value::Float(-f)),
            Value::Boolean(b) => Ok(Value::Boolean(!b)),
            Value::String(_) => Err(SpreadSheetError::new(
                "invalid operation: -string".to_string(),
            )),
            Value::Array(_) => Err(SpreadSheetError::new(
                "invalid operation: -array".to_string(),
            )),
            Value::Object(_) => Err(SpreadSheetError::new(
                "invalid operation: -object".to_string(),
            )),
        }
    }

    pub fn div(&self, rhs: &Value) -> SpreadSheetResult<Value> {
        match self {
            Value::Integer(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Float(*lhs as f64 / *rhs as f64)),
                Value::Float(rhs) => Ok(Value::Float(*lhs as f64 / rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer / string".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer / boolean".to_string(),
                )),
                Value::Array(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer / array".to_string(),
                )),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer / object".to_string(),
                )),
            },
            Value::Float(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Float(lhs / *rhs as f64)),
                Value::Float(rhs) => Ok(Value::Float(lhs / rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: float / string".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: float / boolean".to_string(),
                )),
                Value::Array(_) => Err(SpreadSheetError::new(
                    "invalid operation: float / array".to_string(),
                )),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: float / object".to_string(),
                )),
            },
            Value::String(_) => Err(SpreadSheetError::new(
                "invalid operation: string / _".to_string(),
            )),
            Value::Boolean(_) => Err(SpreadSheetError::new(
                "invalid operation: boolean / _".to_string(),
            )),
            Value::Array(_) => Err(SpreadSheetError::new(
                "invalid operation: array / _".to_string(),
            )),
            Value::Object(_) => Err(SpreadSheetError::new(
                "invalid operation: object / _".to_string(),
            )),
        }
    }

    pub fn mul(&self, rhs: &Value) -> SpreadSheetResult<Value> {
        match self {
            Value::Integer(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Integer(lhs * rhs)),
                Value::Float(rhs) => Ok(Value::Float(*lhs as f64 * rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer * string".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer * boolean".to_string(),
                )),
                Value::Array(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer * array".to_string(),
                )),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: integer * object".to_string(),
                )),
            },
            Value::Float(lhs) => match rhs {
                Value::Integer(rhs) => Ok(Value::Float(lhs * *rhs as f64)),
                Value::Float(rhs) => Ok(Value::Float(lhs * rhs)),
                Value::String(_) => Err(SpreadSheetError::new(
                    "invalid operation: float * string".to_string(),
                )),
                Value::Boolean(_) => Err(SpreadSheetError::new(
                    "invalid operation: float * boolean".to_string(),
                )),
                Value::Array(_) => Err(SpreadSheetError::new(
                    "invalid operation: float * array".to_string(),
                )),
                Value::Object(_) => Err(SpreadSheetError::new(
                    "invalid operation: float * object".to_string(),
                )),
            },
            Value::String(_) => Err(SpreadSheetError::new(
                "invalid operation: string / _".to_string(),
            )),
            Value::Boolean(_) => Err(SpreadSheetError::new(
                "invalid operation: boolean / _".to_string(),
            )),
            Value::Array(_) => Err(SpreadSheetError::new(
                "invalid operation: array / _".to_string(),
            )),
            Value::Object(_) => Err(SpreadSheetError::new(
                "invalid operation: object / _".to_string(),
            )),
        }
    }
}

#[derive(Debug, Default, Clone)]
pub struct Scopes {
    /// The active scope.
    pub top: Scope,
    /// The stack of lower scopes.
    pub scopes: Vec<Scope>,
}

pub trait IntoValue {
    fn into_value(self) -> Value;
}

impl Scopes {
    /// Create a new, empty hierarchy of scopes.
    pub fn new() -> Self {
        Self {
            top: Scope::new(),
            scopes: vec![],
        }
    }

    /// Enter a new scope.
    pub fn enter(&mut self) {
        self.scopes.push(std::mem::take(&mut self.top));
    }

    /// Exit the topmost scope.
    ///
    /// This panics if no scope was entered.
    pub fn exit(&mut self) {
        self.top = self.scopes.pop().expect("no pushed scope");
    }

    /// Try to access a variable immutably.
    pub fn get(&self, var: &str) -> SpreadSheetResult<&Value> {
        std::iter::once(&self.top)
            .chain(self.scopes.iter().rev())
            .find_map(|scope| scope.get(var))
            .ok_or_else(|| unknown_variable(var))
    }

    /// Try to access a variable mutably.
    pub fn get_mut(&mut self, var: &str) -> SpreadSheetResult<&mut Value> {
        std::iter::once(&mut self.top)
            .chain(&mut self.scopes.iter_mut().rev())
            .find_map(|scope| scope.get_mut(var))
            .ok_or_else(|| unknown_variable(var))?
    }

    pub fn resolve_identifier(&self, id: &str) -> Option<&Value> {
        let mut path = PathSplitter::new(&id[1..]);
        if let Some(name) = path.next() {
            if let Ok(value) = self.get(name) {
                value.resolve(&mut path)
            } else {
                None
            }
        } else {
            None
        }
    }

    pub fn resolve(&self, expression: Expression) -> Option<Value> {
        match expression {
            Expression::Value(v) => Some(v),
            Expression::Identifier(id) => self.resolve_identifier(id).cloned(),
        }
    }
}

/// The error message when a variable is not found.
#[cold]
fn unknown_variable(var: &str) -> SpreadSheetError {
    SpreadSheetError::new(format!("unknown variable: {}", var))
}

/// A map from binding names to values.
#[derive(Default, Clone)]
pub struct Scope {
    map: IndexMap<EcoString, Slot>,
}

impl Scope {
    /// Create a new empty scope.
    pub fn new() -> Self {
        Default::default()
    }

    /// Bind a value to a name.
    #[track_caller]
    pub fn define(&mut self, name: impl Into<EcoString>, value: Value) {
        let name = name.into();

        self.map.insert(name, Slot::new(value));
    }

    /// Try to access a variable immutably.
    pub fn get(&self, var: &str) -> Option<&Value> {
        self.map.get(var).map(Slot::read)
    }

    /// Try to access a variable mutably.
    pub fn get_mut(&mut self, var: &str) -> Option<SpreadSheetResult<&mut Value>> {
        self.map.get_mut(var).map(Slot::write)
    }

    /// Iterate over all definitions.
    pub fn iter(&self) -> impl Iterator<Item = (&EcoString, &Value)> {
        self.map.iter().map(|(k, v)| (k, v.read()))
    }
}

impl Debug for Scope {
    fn fmt(&self, f: &mut Formatter) -> fmt::Result {
        f.write_str("Scope ")?;
        f.debug_map()
            .entries(self.map.iter().map(|(k, v)| (k, v.read())))
            .finish()
    }
}

/// A slot where a value is stored.
#[derive(Clone)]
struct Slot {
    /// The stored value.
    value: Value,
}

impl Slot {
    /// Create a new slot.
    fn new(value: Value) -> Self {
        Self { value }
    }

    /// Read the value.
    fn read(&self) -> &Value {
        &self.value
    }

    /// Try to write to the value.
    fn write(&mut self) -> SpreadSheetResult<&mut Value> {
        Ok(&mut self.value)
    }
}

pub struct PathSplitter<'a> {
    pub path: &'a str,
    pub pos: usize,
}

impl<'a> Iterator for PathSplitter<'a> {
    type Item = &'a str;

    fn next(&mut self) -> Option<Self::Item> {
        let start = self.pos;
        let mut end = self.pos;
        while end < self.path.len() {
            match self.path.as_bytes()[end] {
                b'.' => {
                    break;
                }
                _ => {
                    end += 1;
                }
            }
        }
        self.pos = end;
        if start == end {
            None
        } else {
            if self.pos < self.path.len() {
                self.pos += 1;
            }
            Some(&self.path[start..end])
        }
    }
}

impl PathSplitter<'_> {
    pub fn new(path: &str) -> PathSplitter {
        PathSplitter { path, pos: 0 }
    }

    pub fn reset(&mut self) {
        self.pos = 0;
    }
}

impl<'de> Deserialize<'de> for Value {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        deserializer.deserialize_any(ValueVisitor)
    }
}

/// Visitor for value deserialization.
struct ValueVisitor;

impl<'de> Visitor<'de> for ValueVisitor {
    type Value = Value;

    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        formatter.write_str("a spreadsheet-builder value")
    }

    fn visit_bool<E: Error>(self, v: bool) -> Result<Self::Value, E> {
        Ok(Value::Boolean(v))
    }

    fn visit_i8<E: Error>(self, v: i8) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_i16<E: Error>(self, v: i16) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_i32<E: Error>(self, v: i32) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_i64<E: Error>(self, v: i64) -> Result<Self::Value, E> {
        Ok(Value::Integer(v))
    }

    fn visit_u8<E: Error>(self, v: u8) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_u16<E: Error>(self, v: u16) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_u32<E: Error>(self, v: u32) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_u64<E: Error>(self, v: u64) -> Result<Self::Value, E> {
        Ok(Value::Integer(v as i64))
    }

    fn visit_f32<E: Error>(self, v: f32) -> Result<Self::Value, E> {
        Ok(Value::Float(v as f64))
    }

    fn visit_f64<E: Error>(self, v: f64) -> Result<Self::Value, E> {
        Ok(Value::Float(v))
    }

    fn visit_char<E: Error>(self, v: char) -> Result<Self::Value, E> {
        Ok(Value::String(v.to_string()))
    }

    fn visit_str<E: Error>(self, v: &str) -> Result<Self::Value, E> {
        Ok(Value::String(v.to_string()))
    }

    fn visit_borrowed_str<E: Error>(self, v: &'de str) -> Result<Self::Value, E> {
        Ok(Value::String(v.to_string()))
    }

    fn visit_string<E: Error>(self, v: String) -> Result<Self::Value, E> {
        Ok(Value::String(v))
    }

    fn visit_bytes<E: Error>(self, _v: &[u8]) -> Result<Self::Value, E> {
        Err(Error::custom("bytes are not supported"))
    }

    fn visit_borrowed_bytes<E: Error>(self, _v: &'de [u8]) -> Result<Self::Value, E> {
        Err(Error::custom("bytes are not supported"))
    }

    fn visit_byte_buf<E: Error>(self, _v: Vec<u8>) -> Result<Self::Value, E> {
        Err(Error::custom("bytes are not supported"))
    }

    fn visit_none<E: Error>(self) -> Result<Self::Value, E> {
        Err(Error::custom("none is not supported"))
    }

    fn visit_some<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        Value::deserialize(deserializer)
    }

    fn visit_unit<E: Error>(self) -> Result<Self::Value, E> {
        Err(Error::custom("unit is not supported"))
    }

    fn visit_seq<A: SeqAccess<'de>>(self, seq: A) -> Result<Self::Value, A::Error> {
        Ok(Value::Array(Arc::new(Vec::deserialize(
            SeqAccessDeserializer::new(seq),
        )?)))
    }

    fn visit_map<A: MapAccess<'de>>(self, map: A) -> Result<Self::Value, A::Error> {
        Ok(Value::Object(Arc::new(IndexMap::deserialize(
            MapAccessDeserializer::new(map),
        )?)))
    }
}
