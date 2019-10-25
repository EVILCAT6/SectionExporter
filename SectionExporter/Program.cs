using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Moonscraper.ChartParser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SectionExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) ExitProgram("Please drag this program onto a .chart file to open it.");

            string chartPath = args[0];
            string chartFolder = chartPath.Substring(0, chartPath.LastIndexOf('\\'));

            if (!chartPath.EndsWith(".chart")) ExitProgram("You must open a chart file.");
            if (!File.Exists(chartPath)) ExitProgram("File could not be found.");

            Song song = new Song(chartPath);

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(chartFolder + "\\" + song.name + ".xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();

                worksheetPart.Worksheet = new Worksheet(sheetData);

                if(workbookPart.Workbook.Sheets == null)
                {
                    workbookPart.Workbook.AppendChild(new Sheets());
                }

                Sheet sheet = new Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = song.name
                };

                var workingSheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;

                uint rowIndex = 1;
                foreach(Section section in song.sections)
                {
                    Row row = new Row
                    {
                        RowIndex = rowIndex
                    };

                    if (rowIndex == 1) // Header
                    {
                        row.AppendChild(AddCellWithText("Section"));
                        row.AppendChild(AddCellWithText("FC"));
                    }
                    else // Data
                    {
                        row.AppendChild(AddCellWithText(section.title));
                        row.AppendChild(AddCellWithText("No"));
                    }

                    sheetData.AppendChild(row);
                    rowIndex++;
                }

                workbookPart.Workbook.Sheets.AppendChild(sheet);
                workbookPart.Workbook.Save();
                Console.WriteLine("Successfully created Spreadsheet file");
            }
            ExitProgram();
        }

        static void ExitProgram(string reason = "", int code = 0)
        {
            if(reason != "") Console.WriteLine(reason + "\n");

            Console.Write("Press any key to quit the program.");
            Console.ReadKey();
            Environment.Exit(code);
        }

        static Cell AddCellWithText(string text)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.InlineString
            };

            InlineString inlineString = new InlineString();
            Text t = new Text
            {
                Text = text
            };
            inlineString.AppendChild(t);

            cell.AppendChild(inlineString);

            return cell;
        }
    }
}
