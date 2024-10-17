using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

// This class, ExcelDataLoader, is designed to read data from Excel files.
// It provides methods to access sheets, rows, columns, and individual cells.
namespace RAA_Level2
{
    internal class ExcelDataManager
    {
        // This dictionary stores all the data from the Excel workbook
        // Key: Sheet name, Value: List of rows, where each row is a list of cell values
        private Dictionary<string, List<List<string>>> _workbookData;

        // Stores the file path of the Excel file for later use (e.g., when saving changes)
        private string _filePath;

        // Constructor: Initializes the ExcelDataManager with a file path
        public ExcelDataManager(string filePath)
        {
            _filePath = filePath;
            LoadExcelData(filePath);
        }


        // This private method does the actual work of reading the Excel file.
        // It's called by the constructor to populate the _workbookData dictionary.
        private void LoadExcelData(string filePath)
        {
            // Initialize the _workbookData dictionary
            _workbookData = new Dictionary<string, List<List<string>>>();

            // Open the Excel file using the OpenXML library
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                // Get the main part of the workbook
                WorkbookPart workbookPart = document.WorkbookPart;
                // Get all sheets in the workbook
                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>();

                // Loop through each sheet in the workbook
                foreach (Sheet sheet in sheets)
                {
                    // Get the worksheet part for the current sheet
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    // Get the actual data from the worksheet
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Get the name of the current sheet
                    string sheetName = sheet.Name;
                    // Create a list to store the data for the current sheet
                    List<List<string>> currentSheetData = new List<List<string>>();

                    // Loop through each row in the sheet
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        // Create a list of cell values for the current row
                        List<string> rowData = row.Elements<Cell>()
                            .Select(cell => GetCellValue(cell, workbookPart))
                            .ToList();
                        // Add the row data to the current sheet's data
                        currentSheetData.Add(rowData);
                    }

                    // Add the sheet data to the _workbookData dictionary
                    _workbookData[sheetName] = currentSheetData;
                }
            }
        }

        // This private method retrieves the value of a cell.
        // It handles both regular values and shared strings (a way Excel uses to save memory for repeated values).
        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            // Check if the cell contains a shared string
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                // If it's a shared string, we need to look up the actual value
                int ssid = int.Parse(cell.CellValue.Text);
                SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(ssid);
                return ssi.Text.Text;
            }
            // If it's not a shared string, return the cell value (or empty string if null)
            return cell.CellValue?.Text ?? string.Empty;
        }

        // This public method returns a list of all sheet names in the workbook.
        // It's useful when you want to know what sheets are available in the Excel file.
        public IReadOnlyList<string> GetSheetNames()
        {
            // Convert the dictionary keys (sheet names) to a read-only list
            return _workbookData.Keys.ToList().AsReadOnly();
        }

        // This public method retrieves all data from a specific worksheet.
        // It returns a read-only version of the data to prevent modifications.
        public IReadOnlyList<IReadOnlyList<string>> GetWorksheet(string sheetName)
        {
            // Check if the requested sheet exists
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            // Convert the sheet data to a read-only format and return it
            return _workbookData[sheetName].Select(row => (IReadOnlyList<string>)row.AsReadOnly()).ToList().AsReadOnly();
        }

        // This public method retrieves a specific column from a worksheet.
        // It allows you to optionally include or exclude the header row.
        public IReadOnlyList<string> GetColumn(string sheetName, int columnIndex, bool includeHeader = false)
        {
            // Check if the requested sheet exists
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var sheetData = _workbookData[sheetName];
            // Check if the column index is valid
            if (sheetData.Count == 0 || columnIndex < 0 || columnIndex >= sheetData[0].Count)
                throw new ArgumentOutOfRangeException(nameof(columnIndex));

            // Extract the column data
            var columnData = sheetData.Select(row => row.Count > columnIndex ? row[columnIndex] : string.Empty).ToList();

            // Return the column data, optionally skipping the header row
            return includeHeader ? columnData.AsReadOnly() : columnData.Skip(1).ToList().AsReadOnly();
        }

        // This public method retrieves a specific row from a worksheet.
        public IReadOnlyList<string> GetRow(string sheetName, int rowIndex)
        {
            // Check if the requested sheet exists
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var sheetData = _workbookData[sheetName];
            // Check if the row index is valid
            if (rowIndex < 0 || rowIndex >= sheetData.Count)
                throw new ArgumentOutOfRangeException(nameof(rowIndex));

            // Return the requested row as a read-only list
            return sheetData[rowIndex].AsReadOnly();
        }

        // This public method retrieves the value of a specific cell.
        public string GetCellValue(string sheetName, int rowIndex, int columnIndex)
        {
            // Check if the requested sheet exists
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var sheetData = _workbookData[sheetName];
            // Check if the row index is valid
            if (rowIndex < 0 || rowIndex >= sheetData.Count)
                throw new ArgumentOutOfRangeException(nameof(rowIndex));

            var row = sheetData[rowIndex];
            // Check if the column index is valid
            if (columnIndex < 0 || columnIndex >= row.Count)
                throw new ArgumentOutOfRangeException(nameof(columnIndex));

            // Return the value of the specified cell
            return row[columnIndex];
        }      



        public void SaveChanges()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                foreach (var sheetData in _workbookData)
                {
                    string sheetName = sheetData.Key;
                    List<List<string>> sheetRows = sheetData.Value;

                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                    if (sheet != null)
                    {
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                        UpdateSheetData(worksheetPart, sheetRows);
                    }
                }
            }
        }

        private void UpdateSheetData(WorksheetPart worksheetPart, List<List<string>> sheetRows)
        {
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            sheetData.RemoveAllChildren();

            for (int rowIndex = 0; rowIndex < sheetRows.Count; rowIndex++)
            {
                Row row = new Row();
                for (int colIndex = 0; colIndex < sheetRows[rowIndex].Count; colIndex++)
                {
                    Cell cell = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(sheetRows[rowIndex][colIndex])
                    };
                    row.AppendChild(cell);
                }
                sheetData.AppendChild(row);
            }
        }

        // New method to update a specific cell
        public void UpdateCell(string sheetName, int rowIndex, int columnIndex, string value)
        {
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var sheetData = _workbookData[sheetName];
            if (rowIndex < 0 || rowIndex >= sheetData.Count)
                throw new ArgumentOutOfRangeException(nameof(rowIndex));

            var row = sheetData[rowIndex];
            if (columnIndex < 0 || columnIndex >= row.Count)
                throw new ArgumentOutOfRangeException(nameof(columnIndex));

            row[columnIndex] = value;
        }

        // New method to add a new row
        public void AddRow(string sheetName, List<string> rowData)
        {
            if (!_workbookData.ContainsKey(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            _workbookData[sheetName].Add(rowData);
        }

    }
}