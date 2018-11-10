use regex;
use std::error;
use std::fmt::{Display, Formatter, Result};

#[derive(Debug)]
pub enum ExcelErr {
    Xml(String),
    Regex(regex::Error),
}

impl error::Error for ExcelErr {
    fn cause(&self) -> Option<&error::Error> {
        match self {
            ExcelErr::Xml(_) => None,
            ExcelErr::Regex(ref err) => Some(err),
        }
    }
}

impl Display for ExcelErr {
    fn fmt(&self, f: &mut Formatter) -> Result {
        match self {
            ExcelErr::Xml(ref msg) => write!(f, "{}", msg),
            ExcelErr::Regex(ref err) => write!(f, "{}", err),
        }
    }
}

impl From<regex::Error> for ExcelErr {
    fn from(err: regex::Error) -> Self {
        ExcelErr::Regex(err)
    }
}
