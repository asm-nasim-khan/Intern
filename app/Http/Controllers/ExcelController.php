<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelController extends Controller
{
    public function readExcel()
    {
        // Path to the Excel file
        $excelFile = public_path('Students.xlsx');

        // Load the Excel file
        $spreadsheet = IOFactory::load($excelFile);

        // Get the first worksheet in the Excel file
        $worksheet = $spreadsheet->getActiveSheet();

        //
        // Load the Excel file
        // $spreadsheet = IOFactory::load($excelFile);

        // // Set the index of the sheet you want to switch to (0-based index)
        // $sheetIndex = 1; // for example, to switch to the second sheet, use index 1

        // // Set the active sheet by index
        // $spreadsheet->setActiveSheetIndex($sheetIndex);

        // // Get the active sheet
        // $worksheet = $spreadsheet->getActiveSheet();
        //

        // Get the highest row and column in the worksheet
       // Get the highest row and column in the worksheet
        // Get the highest row and column in the worksheet
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = Coordinate::columnIndexFromString($worksheet->getHighestColumn());
        // Iterate through each row in the worksheet
        for ($row = 1; $row <= $highestRow; $row++) {
            // Iterate through each column in the row
            for ($col = 1; $col <= $highestColumn; $col++) {
                // Get the column letter from the column index
                $columnLetter = Coordinate::stringFromColumnIndex($col);
                // Get the value of the current cell
                $cellValue = $worksheet->getCell($columnLetter . $row)->getValue();
                // Output the cell value in a formatted way
                printf("%-30s", $cellValue); // Adjust the width as per your needs
            }
            echo "\n"; // Move to the next line after printing all cells in a row
        }
    }


    public function createExcel()
    {
        // Create a new PhpSpreadsheet instance
        $spreadsheet = new Spreadsheet();

        // Set active sheet
        $spreadsheet->setActiveSheetIndex(0);
        $sheet = $spreadsheet->getActiveSheet();

        // Sample data to write to the Excel file
        $data = [
            [ 'Name', 'Age', 'Country' ],
            [ 'John', 30, 'USA' ],
            [ 'Jane', 25, 'UK' ],
            [ 'Doe', 40, 'Canada' ]
        ];

        // Populate cells with data
        foreach ($data as $row => $rowData) {
            foreach ($rowData as $col => $cellData) {
                $sheet->setCellValueByColumnAndRow($col + 1, $row + 1, $cellData);
            }
        }

        // Create a writer object
        $writer = new Xlsx($spreadsheet);

        // Set the file name and path to save the Excel file
        $excelFilePath = public_path('output.xlsx');

        // Save the spreadsheet to a file
        $writer->save($excelFilePath);

        return response()->download($excelFilePath)->deleteFileAfterSend(true);
    }

}
