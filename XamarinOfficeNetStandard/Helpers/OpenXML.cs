using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Plugin.NetStandardStorage.Abstractions.Interfaces;
using Plugin.NetStandardStorage.Abstractions.Types;
using Plugin.NetStandardStorage.Implementations;

using System.Linq;
using XamarinOfficeNetStandard.Models;
using System.Collections.Generic;

namespace XamarinOfficeNetStandard.Helpers
{
    public class OpenXML
    {
        private IFileSystem _fileSystem = new FileSystem();

        private IFolder OfficeFolder => _fileSystem.LocalStorage.CreateFolder("Office", CreationCollisionOption.OpenIfExists);

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        public string GenerateExcel(string fileName)
        {
            //Fix for https://github.com/OfficeDev/Open-XML-SDK/issues/221
            Environment.SetEnvironmentVariable("MONO_URI_DOTNETRELATIVEORABSOLUTE", "true");

            //Create a new spreadsheet file - remark will overwrite existing ones with the same name!
            string fullPath = Path.Combine(OfficeFolder.FullPath, fileName);
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();

            spreadsheetDocument.Close();

            return fullPath;
        }

        public void InsertDataIntoSheet(string fileName, string sheetName, ExcelData data)
        {
            //Fix for https://github.com/OfficeDev/Open-XML-SDK/issues/221
            Environment.SetEnvironmentVariable("MONO_URI_DOTNETRELATIVEORABSOLUTE", "true");

            // Open the document for editing
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                //Set the sheet name on the first sheet
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault();

                sheet.Name = sheetName;

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                Row headerRow = sheetData.AppendChild(new Row());

                foreach(string header in data.Headers)
                {
                    Cell cell = ConstructCell(header, CellValues.String);
                    headerRow.Append(cell);
                }

                foreach(List<string> dataList in data.Values)
                {
                    Row dataRow = sheetData.AppendChild(new Row());
                    foreach(string dataElement in dataList)
                    {
                        Cell cell = ConstructCell(dataElement, CellValues.String);
                        dataRow.Append(cell);
                    }
                }

                workbookPart.Workbook.Save();
            }
        }

    }
}
