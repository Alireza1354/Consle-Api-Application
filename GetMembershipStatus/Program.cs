using ClosedXML.Excel;

namespace GetMembershipStatus;

public class Program
{
    static async Task Main()
    {
        // ********************** Get file path ************************

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Welcome!");
        Console.WriteLine($"------------------------------------");

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"What is the name of Your parameter?");
        Console.ForegroundColor = ConsoleColor.DarkYellow;
        Console.WriteLine($"1. {nameof(InputParamType.PersonalCode)}    2. {nameof(InputParamType.NationalCode)}");


    repeatReadSelectedChoise:
        if (Enum.TryParse(Console.ReadLine(), out InputParamType selectedChoice))
        {
            switch (selectedChoice)
            {
                case InputParamType.PersonalCode:
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write($"Ok. Your parameter is ---> ");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write($"{selectedChoice}" + Environment.NewLine + Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                case InputParamType.NationalCode:
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write($"Ok. Your parameter is ---> ");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write($"{selectedChoice}" + Environment.NewLine + Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                default:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Please Enter the number 1 or 2");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    goto repeatReadSelectedChoise;
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"The entered number is not valid" + Environment.NewLine);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Please enter 1 or 2");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"1. {nameof(InputParamType.PersonalCode)}    2. {nameof(InputParamType.NationalCode)}");
            goto repeatReadSelectedChoise;
        }

    repeatFilePath:
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Please enter your file path:");
        Console.ForegroundColor = ConsoleColor.White;

    repeatReadFilePath:
        string filePath = Console.ReadLine() ?? "";

        if (string.IsNullOrEmpty(filePath))
        {
            goto repeatFilePath;
        }

        if (File.Exists(filePath))
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write($"OK. Your file path is --->   ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{filePath}" + Environment.NewLine + Environment.NewLine);
            Console.ForegroundColor = ConsoleColor.Yellow;

            var fileInfo = new FileInfo(filePath);
            string fileName = fileInfo.Name;

            Console.Write($"Are you sure to continue with this file ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($" {fileName} ?" + Environment.NewLine);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"1.Yes  2.No Select another file.");
            string result = Console.ReadLine() ?? "";
            switch (result)
            {
                case "1":
                    break;
                case "2":
                    goto repeatFilePath;
                default:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Please enter the number 1 or 2 ");
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"1.Yes  2.No Select another file.");
                    goto repeatReadFilePath;
            }

            //E:\MyData\Book1.xlsx  <--- Sample file path
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Your file does not exist or your file path is invalid!" + Environment.NewLine);
            Console.ForegroundColor = ConsoleColor.White;
            goto repeatFilePath;
        }

    // ***************** Get sheet name ****************************
    enterSheetName:

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(Environment.NewLine + $"Please enter your sheet name:");
        Console.ForegroundColor = ConsoleColor.White;

        string sheetName = Console.ReadLine() ?? "";
        if (string.IsNullOrEmpty(sheetName))
        {
            goto enterSheetName;
        }

        Console.ForegroundColor = ConsoleColor.Blue;
        Console.Write($"OK. Your sheet name is ---> ");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write($" {sheetName}" + Environment.NewLine);
        Console.ForegroundColor = ConsoleColor.White;

        // ***************** Get WorkBook ******************************
        try
        {
            using XLWorkbook workbook = new(filePath);
            if (workbook.TryGetWorksheet(sheetName, out IXLWorksheet xLWorksheet))
            {
                int countColumnUsed = xLWorksheet.ColumnsUsed().Count();

                // Check that the number of columns is not more than one
                if (countColumnUsed != 1)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"There is a problem with your Excel file!");
                    Console.WriteLine($"Please check that the Excel file has only one column." + Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    workbook.Dispose();
                    goto repeatFilePath;
                }

                IXLColumn firstColumnUsed = xLWorksheet.FirstColumnUsed();
                int firstColumNumber = firstColumnUsed.RangeAddress.FirstAddress.ColumnNumber;

                IXLRow firstRowUsed = xLWorksheet.FirstRowUsed();
                int firstRowNumber = firstRowUsed.RangeAddress.FirstAddress.RowNumber;

                IXLRow lastCellUsed = xLWorksheet.LastRowUsed();
                int lastRowUsedNumber = lastCellUsed.RangeAddress.LastAddress.RowNumber;

                // Check to first row that must empty
                if (firstRowNumber == 1)
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.WriteLine($"There is a problem with your Excel file!");
                    Console.WriteLine($"The first row of your Excel file must be empty." + Environment.NewLine);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    workbook.Dispose();
                    goto repeatFilePath;
                }

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine
                    ($"First Row Number : {firstRowNumber} *** " +
                     $"First Column Number: {firstColumNumber}");
                Console.WriteLine("------------------------------------------------" + Environment.NewLine);

                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine("Please wait a moment..." + Environment.NewLine);
                Console.ForegroundColor = ConsoleColor.Yellow;

                //*************************************************************************************************
                ApiResult apiResultDictionary = new();

                XLCellValue DataCellValue;

                int rowUsedIndex = 1;
                int rowNumber = firstRowNumber;
                int colNumber = firstColumNumber;

                int stCode = 0;
                do
                {
                    // Get cell value from workSheet
                    DataCellValue =
                         workbook
                        .Worksheet(name: sheetName)
                        .Cell(row: rowNumber, column: firstColumNumber).Value;

                    if (CellValueIsValid(DataCellValue, selectedChoice))
                    {
                        apiResultDictionary =
                        await Api.GetData(queryParam: DataCellValue.ToString(), inputParamType: selectedChoice);

                        stCode = apiResultDictionary.StatusCode;

                        if (stCode == 201 && apiResultDictionary.JsonData is not null)
                        {
                            foreach (KeyValuePair<string, string> keyValuePair in apiResultDictionary.JsonData)
                            {
                                var key = keyValuePair.Key;
                                var value = keyValuePair.Value;

                                workbook.Worksheet(name: sheetName).Cell(row: firstRowNumber - 1, column: ++colNumber).Value = key;
                                workbook.Worksheet(name: sheetName).Cell(row: rowNumber, column: colNumber).Value = value;
                                workbook.Save();
                            }
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Row: {rowUsedIndex} -- Very Good ---> {DataCellValue}");

                            colNumber = firstColumNumber;
                        }
                        else
                        {
                            if (stCode == 400)
                            {
                                workbook
                               .Worksheet(name: sheetName)
                               .Cell(row: rowNumber, column: firstColumNumber + 1).Value = "Not found";
                            }
                            //workbook
                            //      .Worksheet(name: sheetName)
                            //      .Cell(row: rowNumber, column: firstColumNumber + 1).Style.Fill.BackgroundColor = XLColor.Red;
                            workbook.Save();
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Row: {rowUsedIndex} --> Error: " +
                                $"{Environment.NewLine}{apiResultDictionary.ResponseMessage}");
                            Console.ForegroundColor = ConsoleColor.White;

                            if (stCode == -1) Environment.Exit(0);
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Row: {rowUsedIndex} -- Is Not valid ---> {DataCellValue}");
                        workbook
                            .Worksheet(name: sheetName)
                            .Cell(row: rowNumber, column: firstColumNumber + 1).Value = "Is not valid";
                        workbook.Save();
                    }

                    rowNumber++;
                    rowUsedIndex++;

                } while (stCode != 201 && rowNumber <= lastRowUsedNumber);

                int remainRows = lastRowUsedNumber - rowNumber;

                for (int i = 0; i <= remainRows; i++)
                {
                    // Get cell value from workSheet
                    DataCellValue =
                         workbook
                        .Worksheet(name: sheetName)
                        .Cell(row: rowNumber, column: firstColumNumber).Value;

                    if (CellValueIsValid(DataCellValue, selectedChoice))
                    {
                        apiResultDictionary =
                        await Api.GetData(queryParam: DataCellValue.ToString(), inputParamType: selectedChoice);

                        stCode = apiResultDictionary.StatusCode;

                        if (stCode == 201 && apiResultDictionary.JsonData is not null)
                        {
                            colNumber = firstColumNumber;

                            foreach (KeyValuePair<string, string> keyValuePair in apiResultDictionary.JsonData)
                            {
                                var key = keyValuePair.Key;
                                var value = keyValuePair.Value;

                                workbook.Worksheet(name: sheetName).Cell(row: rowNumber, column: ++colNumber).Value = value;
                            }

                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Row: {rowUsedIndex} -- Very Good ---> {DataCellValue}");
                        }
                        else
                        {
                            workbook
                                 .Worksheet(name: sheetName)
                                 .Cell(row: rowNumber, column: firstColumNumber + 1).Value = apiResultDictionary.ResponseMessage;

                            workbook
                                 .Worksheet(name: sheetName)
                                 .Cell(row: rowNumber, column: firstColumNumber + 1).Style.Fill.BackgroundColor = XLColor.Red;

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Row: {rowUsedIndex} -- Status Code: {apiResultDictionary.StatusCode}");
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Row: {rowUsedIndex} -- Not Valid ---> {DataCellValue}");
                    }
                    rowNumber++;
                    rowUsedIndex++;
                }

                workbook.Save();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"***Finished******");
                Console.WriteLine($"Press Enter to exit.");
                Console.ForegroundColor = ConsoleColor.White;
                Console.ReadLine();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Your excel file or sheet name does not exist !!!" + Environment.NewLine);
                Console.WriteLine($"Look at here!");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Please Enter 1 or 2 ---> ");
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"1.Enter new file path   2.Enter new sheet name" + Environment.NewLine);
                Console.ForegroundColor = ConsoleColor.DarkYellow;
            reEnterFileOrSheetName:
                string reSelectFile = Console.ReadLine() ?? "";
                switch (reSelectFile)
                {
                    case "1":
                        workbook.Dispose();
                        goto repeatFilePath;
                    case "2":
                        workbook.Dispose();
                        goto enterSheetName;
                    default:
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Please Enter 1 or 2 ---> ");
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine($"1.Enter new file path   2.Enter new sheet name" + Environment.NewLine);
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        goto reEnterFileOrSheetName;
                }
            }
        }
        catch (Exception ex)
        {
            string msg = ex.Message;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(msg);
            Console.WriteLine(Environment.NewLine);
            goto repeatFilePath;
        }
    }

    public static bool CellValueIsValid(XLCellValue cellValue, InputParamType selectedChoice)
    {
        if (cellValue.IsBlank) { return false; }

        if (String.IsNullOrWhiteSpace(cellValue.ToString())) { return false; }

        if (!cellValue.ToString().All(char.IsDigit)) { return false; }

        if (selectedChoice == InputParamType.NationalCode && cellValue.ToString().Length == 10)
        {
            return true;
        }

        if (selectedChoice == InputParamType.PersonalCode && cellValue.ToString().Length == 8)
        {
            return true;
        }

        return false;
    }

    public enum InputParamType
    {
        PersonalCode = 1,
        NationalCode = 2
    }
}