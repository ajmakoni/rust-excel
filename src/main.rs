use rust_xlsxwriter::{Workbook, XlsxError};

#[tokio::main]
async fn main() {
    println!("Server started at localhost:8080");
    match doexcel().await {
        Ok(_ok) => {
            println!("Excel document successfully generated");
        },
        Err(e) => {
            println!("Error generating excel document: {}", e);
        }
    }
    // let _ = doexcel().await;
}

async fn doexcel() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();
    let titles = vec![
        "First Name",
        "Last Name",
        "Contact Number",
        "Address",
        "Email",
        "Id",
    ];
    let mut cn = 0;
    for title in titles {
        worksheet.write(0, cn, title)?;
        cn += 1;
    }

    let details = vec![
        "Mophius",
        "Dynamic",
        "0100987889",
        "Home address, Corner Palace",
        "email",
        "12345",
    ];

    for a in 1..=5000 {
        let mut pos = 0;
        for records in &details {
            let mut entry = format!("{}{}", records, a);
            if pos == 4 {
                entry = format!("test{}@r8code.dev", a);
            }
            worksheet.write(a, pos, entry)?;
            pos += 1;
        }
    }

    // Write a string to cell (0, 0) = A1.
    //worksheet.write(0, 0, "Hello")?;

    // Write a number to cell (1, 0) = A2.

    // Save the file to disk.
    workbook.save("hello.xlsx")?;

    Ok(())
}

#[cfg(test)]
mod tests{
    use crate::doexcel;

    #[tokio::test]
    async fn test_generation_of_excel(){
        match doexcel().await {
            Ok(_)=>{
                println!("Successful")
            },
            Err(e)=>{
                println!("Failed: {:?}",e);
            }
        }
    }
}