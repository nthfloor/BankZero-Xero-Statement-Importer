import * as XLSX from 'xlsx';

import { createObjectCsvWriter as createCsvWriter } from 'csv-writer';
import * as fs from 'fs/promises';
import * as path from 'path';

// Define the paths to the files
const templateCsvPath = './StatementImportTemplate.en-US.csv';

// Function to convert epoch to date string
function convertExcelSerialDateToDateStr(serialDate: number): string {
    // Create a base date (January 0, 1900)
    const baseDate = new Date(Date.UTC(1900, 0, -1));

    // Add the serial date number to the base date
    const resultDate = new Date(baseDate.getTime() + serialDate * 24 * 60 * 60 * 1000);

    // Format the date as a string (e.g., YYYY-MM-DD)
    const dateString = resultDate.toISOString().split('T')[0];

    return dateString;
}

function compileTransactionDescription(transaction: any): {
    Payee: string;
    Description: string;
    Reference: string;
} {
    switch (transaction['Type']) {
        case 'Transfer out':
            return {
                Payee: transaction['Description 1'],
                Description: `Bank Transfer to ${transaction['Description 1']}`,
                Reference: transaction['Description 2'],
            }
        case 'Transfer in':
            return {
                Payee: transaction['Description 1'],
                Description: `Bank Transfer from ${transaction['Description 1']}`,
                Reference: transaction['Description 2'],
            }
        case 'Interest received':
            return {
                Payee: 'Bank Zero',
                Description: transaction['Description 1'],
                Reference: transaction['Description 2'],
            }
        default:
            throw Error('Invalid state');
    }
}

async function transformAndExportData(excelFilePath: string) {
    try {
        // Read the template CSV to understand the required format
        const data = await fs.readFile(templateCsvPath, 'utf8');
        const headers = data.split('\n')[0].split(',');
        console.log('The CSV template file has been read successfully.');

        // Read the Excel file
        const workbook = XLSX.readFile(excelFilePath);
        const sheetNameList = workbook.SheetNames;
        const xlsData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
        console.log('The Excel file has been read successfully: ', excelFilePath);

        // Transform the data to match the structure of the template CSV
        const transformedData = xlsData.map((row: any) => {
            // Transform each row here based on your specific requirements
            // This is an example transformation
            const XeroTransactionData = compileTransactionDescription(row);
            return {
                [headers[0]]: convertExcelSerialDateToDateStr(row[headers[0]]), // Map the Excel data to the Date Field
                [headers[1]]: row[headers[1]], // Map the Excel data to the Amount Field
                [headers[2]]: XeroTransactionData.Payee, // Map the Excel data to the Payee Field
                [headers[3]]: XeroTransactionData.Description, // Map the Excel data to the Description Field
                [headers[4]]: XeroTransactionData.Reference, // Map the Excel data to the Reference Field
            };
        });

        // Define the CSV Writer with the headers from the template file
        const outputCsvPath = `./${excelFilePath}_TransformedData.csv`;
        const csvWriter = createCsvWriter({
            path: outputCsvPath,
            header: headers.map((headerName) => ({ id: headerName, title: headerName })),
        });

        // Write the transformed data to a new CSV file
        await csvWriter.writeRecords(transformedData);
        console.log('The CSV file was written successfully: ', outputCsvPath);

    } catch (err) {
        console.error("An error occurred:", err);
    }
}

async function transformAllBankZeroStatementFiles() {
    try {
        // const excelFilePath = './Optimal ALS_Savings_December2023.xls';
        const directoryPath = './'

        // Read the directory
        const files = await fs.readdir(directoryPath, { withFileTypes: true });
        // Filter excel files that end with ".xls"
        const xlsFiles = files
            .filter(file => 
                file.isFile() 
                && (file.name.startsWith('Icecream') || file.name.startsWith('Cash reserves') || file.name.startsWith('Optimal')) 
                && path.extname(file.name).toLowerCase() === '.xls')
            .map(file => file.name);
            
        for(const excelFilePath of xlsFiles) {
            await transformAndExportData(excelFilePath);
        }
        console.log(`Transformed ${xlsFiles.length} files successfully.`);
    } catch (error) {
        console.error('An error occurred:', error);
        return [];
    }
}

transformAllBankZeroStatementFiles();
