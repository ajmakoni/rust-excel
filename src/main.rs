use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, XlsxError};

// ================================================================
// Path to the Excel file you saved from the browser.
// Change this to wherever you saved it.
// ================================================================
const INPUT_XLSX_PATH: &str = "/home/user/Downloads/students.xlsx";
const OUTPUT_XLSX_PATH: &str = "hello.xlsx";
const SHEET_NUMBER: usize = 0;

fn main() {
    println!("Program has started");

    match doexcel()/* doexcelfromlocalpath(INPUT_XLSX_PATH, SHEET_NUMBER) */ {
        Ok(count) => {
            println!(
                "Excel document successfully generated ({} records)",
                count
            );
        }
        Err(e) => {
            println!("Error generating excel document: {}", e);
        }
    }
}


// ================================================================
// YOUR EXISTING FUNCTION - UNCHANGED
// ================================================================

fn doexcel() -> Result<i32, XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let titles = vec![
        "First Name",
        "Last Name",
        "Contact Number",
        "Address",
        "Email",
        "Id",
    ];

    for (cn, title) in titles.iter().enumerate() {
        worksheet.write(0, cn as u16, *title)?;
    }

    let details = vec![
        "Mophius",
        "Dynamic",
        "0100987889",
        "Home address, Corner Palace",
        "email",
        "12345",
    ];

    for a in 1..=50u32 {
        for (pos, records) in details.iter().enumerate() {
            let entry = if pos == 4 {
                format!("nt{}@r8code.dev", a)
            } else {
                format!("{}{}", records, a)
            };
            worksheet.write(a, pos as u16, entry)?;
        }
    }

    workbook.save("hello.xlsx")?;
    Ok(50)
}


// ================================================================
// NEW: read a local .xlsx, transform, write hello.xlsx
// ================================================================

fn doexcelfromlocalpath(
    input_path: &str,
    sheet_number: usize,
) -> Result<u32, XlsxError> {
    println!("Reading local file: {}", input_path);

    let mut source: Xlsx<_> = open_workbook(input_path)
        .map_err(|e| {
            XlsxError::ParameterError(format!(
                "Failed to open '{}': {}",
                input_path, e
            ))
        })?;

    let sheet_names = source.sheet_names().to_owned();
    println!("Available sheets:");
    for (i, name) in sheet_names.iter().enumerate() {
        println!("  {}: {}", i, name);
    }

    if sheet_number >= sheet_names.len() {
        return Err(XlsxError::ParameterError(format!(
            "Sheet {} does not exist. Available: {:?}",
            sheet_number, sheet_names
        )));
    }
    println!(
        "Using sheet {}: {}",
        sheet_number, sheet_names[sheet_number]
    );

    let range = source
        .worksheet_range_at(sheet_number)
        .ok_or_else(|| {
            XlsxError::ParameterError(format!(
                "Unable to read sheet {}",
                sheet_number
            ))
        })?
        .map_err(|e| {
            XlsxError::ParameterError(format!(
                "Failed to read worksheet: {}",
                e
            ))
        })?;

    // ------------------------------------------------------------
    // Find header columns
    // ------------------------------------------------------------
    let header = range.rows().next().ok_or_else(|| {
        XlsxError::ParameterError("Excel sheet is empty".to_string())
    })?;

    let mut name_col = None;
    let mut surname_col = None;
    let mut student_number_col = None;
    let mut email_col = None;

    for (index, cell) in header.iter().enumerate() {
        let value = cell.to_string().trim().to_uppercase();
        match value.as_str() {
            "NAME" => name_col = Some(index),
            "SURNAME" => surname_col = Some(index),
            "STUDENT NUMBER" => student_number_col = Some(index),
            "EMAIL" => email_col = Some(index),
            _ => {}
        }
    }

    let name_col = name_col.ok_or_else(|| {
        XlsxError::ParameterError("NAME column not found".into())
    })?;
    let surname_col = surname_col.ok_or_else(|| {
        XlsxError::ParameterError("SURNAME column not found".into())
    })?;
    let student_number_col = student_number_col.ok_or_else(|| {
        XlsxError::ParameterError("STUDENT NUMBER column not found".into())
    })?;
    let email_col = email_col.ok_or_else(|| {
        XlsxError::ParameterError("EMAIL column not found".into())
    })?;

    // ------------------------------------------------------------
    // Create output workbook
    // ------------------------------------------------------------
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for (cn, title) in [
        "First Name",
        "Last Name",
        "Contact Number",
        "Address",
        "Email",
        "Id",
    ]
    .iter()
    .enumerate()
    {
        worksheet.write(0, cn as u16, *title)?;
    }

    // ------------------------------------------------------------
    // Copy / transform rows
    // ------------------------------------------------------------
    let mut output_row = 1u32;

    for row in range.rows().skip(1) {
        let mut first_name = row
            .get(name_col)
            .map(|v| v.to_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        let mut last_name = row
            .get(surname_col)
            .map(|v| v.to_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        let mut student_number = row
            .get(student_number_col)
            .map(|v| v.to_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        let email = row
            .get(email_col)
            .map(|v| v.to_string())
            .unwrap_or_default()
            .trim()
            .to_string();

        // Skip fully empty rows
        if first_name.is_empty()
            && last_name.is_empty()
            && student_number.is_empty()
            && email.is_empty()
        {
            continue;
        }

        // --------------------------------------------------------
        // Fallbacks: if ID / name is empty, use the local part
        // of the email (everything before '@').
        // --------------------------------------------------------
        let email_local_part: String = email
            .split('@')
            .next()
            .unwrap_or("")
            .to_string();

        if student_number.is_empty() && !email_local_part.is_empty() {
            student_number = email_local_part.clone();
        }

        if first_name.is_empty() && !email_local_part.is_empty() {
            first_name = email_local_part.clone();
        }

        if last_name.is_empty() && !email_local_part.is_empty() {
            last_name = email_local_part.clone();
        }

        worksheet.write(output_row, 0, first_name)?;
        worksheet.write(output_row, 1, last_name)?;
        worksheet.write(output_row, 2, "")?;
        worksheet.write(output_row, 3, "")?;
        worksheet.write(output_row, 4, email)?;
        worksheet.write(output_row, 5, student_number)?;

        output_row += 1;
    }

    workbook.save(OUTPUT_XLSX_PATH)?;

    let count = output_row - 1;
    println!(
        "Saved {} records to {}",
        count, OUTPUT_XLSX_PATH
    );

    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_local_file_generation() {
        match doexcelfromlocalpath(INPUT_XLSX_PATH, SHEET_NUMBER) {
            Ok(n) => println!("Successful, {} records", n),
            Err(e) => println!("Failed: {:?}", e),
        }
    }
}