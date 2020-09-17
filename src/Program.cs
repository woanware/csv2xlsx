using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace csv2xlsx
{
    class Program
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Console.WriteLine(args[0]);
            string fileName = args[0];
            if (!File.Exists(fileName))
            {
                throw new FileNotFoundException(String.Format("File {0} was not found!", fileName));
            }

            string resultPath = Path.GetDirectoryName(args[0]);
            if (resultPath.Length == 0)
            {
                resultPath = ".";
            }
            string resultFile = Path.GetFileNameWithoutExtension(args[0]);

            CreateWorkBook(args[0], Path.Combine(resultPath, resultFile + ".xlsx"));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="csvFilePath"></param>
        /// <param name="xlsxFilePath"></param>
        private static void CreateWorkBook(string csvFilePath, string xlsxFilePath)
        {
            using (var workbook = SpreadsheetDocument.Create(xlsxFilePath, SpreadsheetDocumentType.Workbook))
            {
                List<OpenXmlAttribute> attributeList;

                workbook.AddWorkbookPart();
                WorksheetPart workSheet = workbook.WorkbookPart.AddNewPart<WorksheetPart>();

                using (OpenXmlWriter writer = OpenXmlWriter.Create(workSheet))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    using (FileStream fsRead = new FileStream(csvFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (StreamReader sr = new StreamReader(fsRead))
                    using (CsvReader csvReader = new CsvReader(sr, CultureInfo.InvariantCulture))
                    {
                        string[] values;
                        long row = 1;
                        while (csvReader.Read())
                        {
                            attributeList = new List<OpenXmlAttribute>();
                            // this is the row index
                            attributeList.Add(new OpenXmlAttribute("r", null, row.ToString()));

                            writer.WriteStartElement(new Row(), attributeList);

                            // Get the CSV record
                            values = csvReader.Context.Record;

                            for (long col = 1; col < values.Count(); col++)
                            {
                                attributeList = new List<OpenXmlAttribute>();
                                // this is the data type ("t"), with CellValues.String ("str")
                                attributeList.Add(new OpenXmlAttribute("t", null, "str"));

                                // Write out the cell
                                writer.WriteStartElement(new Cell(), attributeList);
                                writer.WriteElement(new CellValue(values[col].ToString()));
                                writer.WriteEndElement();
                            }

                            writer.WriteEndElement(); // Row
                            row++;
                        }
                    }

                    writer.WriteEndElement(); // SheetData
                    writer.WriteEndElement(); // Worksheet
                }

                CreateWorkSheet(workbook, workSheet);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="workSheet"></param>
        private static void CreateWorkSheet(SpreadsheetDocument workbook, WorksheetPart workSheet)
        {
            using (OpenXmlWriter writer = OpenXmlWriter.Create(workbook.WorkbookPart))
            {
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                writer.WriteElement(new Sheet()
                {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = workbook.WorkbookPart.GetIdOfPart(workSheet)
                });

                writer.WriteEndElement(); // WorkSheet
                writer.WriteEndElement(); // WorkBook
            }
        }
    }
}
