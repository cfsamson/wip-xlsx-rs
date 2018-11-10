#![allow(dead_code)]
use crate::error::ExcelErr;
use regex;
use regex::Regex;
use std::{char, u8};

// TODO: use lazy static to compile regex here
const RANGE_PARTS: &str = r#"(\$?)([A-Z]{1,3})(\$?)(\d+)"#;

pub struct Utility {
    range_parts: regex::Regex,
    col_names: Vec<String>,
}

impl Utility {
    pub fn new() -> Result<Self, ExcelErr> {
        let util = Utility {
            range_parts: regex::Regex::new(r#"(\$?)([A-Z]{1,3})(\$?)(\d+)"#)?,
            col_names: vec![],
        };
        Ok(util)
    }
}

/// Convert a zero indexed row and column cell reference to a A1 style string.
/// row:     The cell row.
/// col:     The cell column.
/// row_abs: Optional flag to make the row absolute.
/// col_abs: Optional flag to make the column absolute.
/// returns: A1 style string.
/// 
/// Panics if index is out of range
fn xl_rowcol_to_cell(
    mut row: usize,
    col: usize,
    row_abs: bool,
    col_abs: bool,
) -> Result<String, ExcelErr> {

    row += 1;
    let col_str = xl_col_to_name(col, col_abs)?;

    let row_col = if row_abs {
        format!("{}{}{}", col_str, "$", row)
    } else {
        format!("{}{}", col_str, row)
    };

    Ok(row_col)
}

fn xl_col_to_name(mut col_num: usize, col_abs: bool) -> Result<String, ExcelErr> {
    col_num += 1;
    let mut col_str: Vec<String> = vec![];
    let col_abs = if col_abs { Some("$") } else { None };

    while col_num != 0 {
        // set remainder from 1 .. 26
        let mut remainder: u8 = (col_num % 26) as u8;
        if remainder == 0 {
            remainder = 26;
        }

        // Convert the remainder to a character.
        let col_letter: char = char::from_u32(('A' as u8 + remainder - 1).into())
            .ok_or_else(|| ExcelErr::Xml("column index out of bounds".to_owned()))?;

        // Accumulate the column letters, in an array. This will accumulate in
        // the "wrong" order by default som we need to switch later
        col_str.push(format!("{}", col_letter));

        col_num = ((col_num - 1) / 26) as usize;
    }

    // if the column is absolute, prepend it with a "$"
    if let Some(s) = col_abs {
        col_str.push(s.to_owned());
    };

    // we reverse them into the right order
    col_str.reverse();
    let res: String = col_str.concat();

    Ok(res)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xl_cell_to_rowcol() {
        let tests = [
            // row, col, A1 string
            (0, 0, "A1"),
            (0, 1, "B1"),
            (0, 2, "C1"),
            (0, 9, "J1"),
            (1, 0, "A2"),
            (2, 0, "A3"),
            (9, 0, "A10"),
            (1, 24, "Y2"),
            (7, 25, "Z8"),
            (9, 26, "AA10"),
            (1, 254, "IU2"),
            (1, 255, "IV2"),
            (1, 256, "IW2"),
            (0, 16383, "XFD1"),
            (1_048_576, 16384, "XFE1048577"),
        ];

        for (row, col, string) in tests.iter() {
            let col = *col as usize;
            let row = *row as usize;
            let exp = *string;
            let got = xl_rowcol_to_cell(row, col, false, false).unwrap();
            assert_eq!(got, exp);
        }
    }

    #[test]
    fn test_xl_rowcol_to_cell_abs() {
        let tests = [
            // row, col, row_abs, col_abs, A1 string
            (0, 0, false, false, "A1"),
            (0, 0, true, false, "A$1"),
            (0, 0, false, true, "$A1"),
            (0, 0, true, true, "$A$1"),
        ];

        for (row, col, row_abs, col_abs, string) in tests.iter() {
            let col = *col as usize;
            let row = *row as usize;
            let exp = *string;
            let got = xl_rowcol_to_cell(row, col, *row_abs, *col_abs).unwrap();
            assert_eq!(got, exp);
        }
    }

    #[test]
    fn test_xl_col_to_name() {
        let tests = [
            // col,  col string
            (0, "A"),
            (1, "B"),
            (2, "C"),
            (9, "J"),
            (24, "Y"),
            (25, "Z"),
            (26, "AA"),
            (254, "IU"),
            (255, "IV"),
            (256, "IW"),
            (16383, "XFD"),
            (16384, "XFE"),
        ];

        for (col, string) in tests.iter() {
            let exp = *string;
            let got = xl_col_to_name(*col as usize, false).unwrap();
            assert_eq!(got, exp);
        }
    }

    #[test]
    fn test_xl_col_to_name_abs() {
        let tests = [
            // col,  col string
            (0, "$A"),
            (25, "$Z"),
            (26, "$AA"),
            (16383, "$XFD"),
        ];

        for (col, string) in tests.iter() {
            let exp = *string;
            let got = xl_col_to_name(*col as usize, true).unwrap();
            assert_eq!(got, exp);
        }
    }
}
