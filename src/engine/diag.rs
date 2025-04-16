use thiserror::Error;

#[derive(Error, Debug)]
pub enum SpreadSheetError {
    #[error("Message: {0}")]
    Message(String),
}

/// A result type with a string error message and hints.
pub type SpreadSheetResult<T> = Result<T, SpreadSheetError>;

impl SpreadSheetError {
    /// Creates a new error with the given message.
    pub fn new(message: String) -> Self {
        Self::Message(message)
    }
}

impl<S> From<S> for SpreadSheetError
where
    S: Into<String>,
{
    fn from(value: S) -> Self {
        Self::new(value.into())
    }
}
