use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, XlsxError};

// ================================================================
// Configuration
// ================================================================

const INPUT_XLSX_PATH: &str = "/home/user/Downloads/students.xlsx";
const OUTPUT_XLSX_PATH: &str = "hello.xlsx";

//  0  = first sheet
//  1  = second sheet
// -1  = merge ALL sheets
const SHEET_NUMBER: i32 = -1;

fn main() {
    println!("Program has started");

    match doexcelfromlocalpath(INPUT_XLSX_PATH, SHEET_NUMBER) {
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
// Existing test/example function
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

    for a in 1..=10u32 {
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
// Read local XLSX
//
// sheet_number:
//   >= 0 = process only that sheet
//   -1   = process all sheets
// ================================================================

fn doexcelfromlocalpath(
    input_path: &str,
    sheet_number: i32,
) -> Result<u32, XlsxError> {
    println!("Reading local file: {}", input_path);

    let mut source: Xlsx<_> = open_workbook(input_path)
        .map_err(|e| {
            XlsxError::ParameterError(format!(
                "Failed to open '{}': {}",
                input_path, e
            ))
        })?;

    // ------------------------------------------------------------
    // Get sheet names
    // ------------------------------------------------------------

    let sheet_names = source.sheet_names().to_owned();

    if sheet_names.is_empty() {
        return Err(XlsxError::ParameterError(
            "Excel workbook contains no sheets".to_string(),
        ));
    }

    println!("Available sheets:");

    for (i, name) in sheet_names.iter().enumerate() {
        println!("  {}: {}", i, name);
    }

    // ------------------------------------------------------------
    // Determine which sheets to process
    // ------------------------------------------------------------

    let sheets_to_process: Vec<usize> = if sheet_number == -1 {
        println!("Sheet number is -1: merging ALL sheets");

        (0..sheet_names.len()).collect()
    } else if sheet_number >= 0 {
        let index = sheet_number as usize;

        if index >= sheet_names.len() {
            return Err(XlsxError::ParameterError(format!(
                "Sheet {} does not exist. Available sheets: {:?}",
                index,
                sheet_names
            )));
        }

        println!(
            "Using sheet {}: {}",
            index,
            sheet_names[index]
        );

        vec![index]
    } else {
        return Err(XlsxError::ParameterError(format!(
            "Invalid sheet number {}. Use -1 for all sheets or >= 0 for a specific sheet.",
            sheet_number
        )));
    };

    // ------------------------------------------------------------
    // Create output workbook
    // ------------------------------------------------------------

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // ------------------------------------------------------------
    // Output header
    // ------------------------------------------------------------

    let titles = [
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

    // First data row.
    let mut output_row = 1u32;

    // Total records processed.
    let mut total_count = 0u32;

    // ------------------------------------------------------------
    // Process each selected sheet
    // ------------------------------------------------------------

    for sheet_index in sheets_to_process {
        let sheet_name = &sheet_names[sheet_index];

        println!();
        println!(
            "============================================================"
        );
        println!(
            "Processing sheet {}: {}",
            sheet_index,
            sheet_name
        );
        println!(
            "============================================================"
        );

        let range = source
            .worksheet_range_at(sheet_index)
            .ok_or_else(|| {
                XlsxError::ParameterError(format!(
                    "Unable to read sheet {} ({})",
                    sheet_index,
                    sheet_name
                ))
            })?
            .map_err(|e| {
                XlsxError::ParameterError(format!(
                    "Failed to read worksheet {} ({}): {}",
                    sheet_index,
                    sheet_name,
                    e
                ))
            })?;

        // --------------------------------------------------------
        // Empty sheet
        // --------------------------------------------------------

        let mut rows = range.rows();

        let header = match rows.next() {
            Some(header) => header,
            None => {
                println!(
                    "Sheet '{}' is empty. Skipping.",
                    sheet_name
                );

                continue;
            }
        };

        // --------------------------------------------------------
        // Find required columns
        // --------------------------------------------------------

        let mut name_col = None;
        let mut surname_col = None;
        let mut student_number_col = None;
        let mut email_col = None;

        for (index, cell) in header.iter().enumerate() {
            let value = cell.to_string().trim().to_uppercase();

            match value.as_str() {
                "NAME" => {
                    name_col = Some(index);
                }

                "SURNAME" => {
                    surname_col = Some(index);
                }

                "STUDENT NUMBER" => {
                    student_number_col = Some(index);
                }

                "EMAIL" => {
                    email_col = Some(index);
                }

                _ => {}
            }
        }

        // --------------------------------------------------------
        // Validate columns
        // --------------------------------------------------------

        let name_col = match name_col {
            Some(col) => col,
            None => {
                return Err(XlsxError::ParameterError(format!(
                    "NAME column not found in sheet '{}'",
                    sheet_name
                )));
            }
        };

        let surname_col = match surname_col {
            Some(col) => col,
            None => {
                return Err(XlsxError::ParameterError(format!(
                    "SURNAME column not found in sheet '{}'",
                    sheet_name
                )));
            }
        };

        let student_number_col = match student_number_col {
            Some(col) => col,
            None => {
                return Err(XlsxError::ParameterError(format!(
                    "STUDENT NUMBER column not found in sheet '{}'",
                    sheet_name
                )));
            }
        };

        let email_col = match email_col {
            Some(col) => col,
            None => {
                return Err(XlsxError::ParameterError(format!(
                    "EMAIL column not found in sheet '{}'",
                    sheet_name
                )));
            }
        };

        println!(
            "Columns found: NAME={}, SURNAME={}, STUDENT NUMBER={}, EMAIL={}",
            name_col,
            surname_col,
            student_number_col,
            email_col
        );

        // --------------------------------------------------------
        // Process rows
        // --------------------------------------------------------

        let mut sheet_count = 0u32;

        for row in rows {
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

            // ----------------------------------------------------
            // Skip completely empty rows
            // ----------------------------------------------------

            if first_name.is_empty()
                && last_name.is_empty()
                && student_number.is_empty()
                && email.is_empty()
            {
                continue;
            }

            // ----------------------------------------------------
            // Email local part
            //
            // john.smith@example.com
            //        ↓
            // john.smith
            // ----------------------------------------------------

            let email_local_part = email
                .split('@')
                .next()
                .unwrap_or("")
                .trim()
                .to_string();

            // ----------------------------------------------------
            // Fallbacks
            // ----------------------------------------------------

            if student_number.is_empty()
                && !email_local_part.is_empty()
            {
                student_number = email_local_part.clone();
            }

            if first_name.is_empty()
                && !email_local_part.is_empty()
            {
                first_name = email_local_part.clone();
            }

            if last_name.is_empty()
                && !email_local_part.is_empty()
            {
                last_name = email_local_part.clone();
            }

            // ----------------------------------------------------
            // Write output row
            // ----------------------------------------------------

            worksheet.write(output_row, 0, first_name)?;
            worksheet.write(output_row, 1, last_name)?;
            worksheet.write(output_row, 2, "")?;
            worksheet.write(output_row, 3, "")?;
            worksheet.write(output_row, 4, email)?;
            worksheet.write(output_row, 5, student_number)?;

            output_row += 1;
            sheet_count += 1;
            total_count += 1;
        }

        println!(
            "Sheet '{}' contributed {} records",
            sheet_name,
            sheet_count
        );
    }

    // ------------------------------------------------------------
    // Save output
    // ------------------------------------------------------------

    workbook.save(OUTPUT_XLSX_PATH)?;

    println!();
    println!(
        "============================================================"
    );
    println!(
        "Saved {} records to {}",
        total_count,
        OUTPUT_XLSX_PATH
    );
    println!(
        "============================================================"
    );

    Ok(total_count)
}


// ================================================================
// Tests
// ================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_local_file_generation() {
        match doexcelfromlocalpath(
            INPUT_XLSX_PATH,
            SHEET_NUMBER,
        ) {
            Ok(n) => println!(
                "Successful, {} records",
                n
            ),

            Err(e) => println!(
                "Failed: {:?}",
                e
            ),
        }
    }
}