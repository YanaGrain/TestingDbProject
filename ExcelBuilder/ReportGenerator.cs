using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;

namespace ExcelBuilder
{
    public class ReportGenerator
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public const int HeaderRowNumber = 1; // headers in the first row of Excel document
        public const int StartRowNumber = 2; // main content begins from 2 row

        private Container _container;

        public ReportGenerator(IBuilder builder)
        {
            _container = builder.Build();
        }

        public void Generate()
        {
            string filePath = @"e:\SolbegSoft\Excel\Report.xlsx";

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                document.CompressionOption = CompressionOption.SuperFast;
                document.AddWorkbookPart();
                WorksheetPart worksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    try
                    {
                        WriteRow(_container.Headers, writer, HeaderRowNumber);
                        WriteRows(_container.Rows, writer, StartRowNumber);
                    }
                    catch (Exception e)
                    {
                        logger.Error(e.Message);
                    }

                    // this is for SheetData
                    writer.WriteEndElement();
                    // this is for Worksheet
                    writer.WriteEndElement();
                    writer.Close();

                    using (var writer1 = OpenXmlWriter.Create(document.WorkbookPart))
                    {
                        writer1.WriteStartElement(new Workbook());
                        writer1.WriteStartElement(new Sheets());

                        writer1.WriteElement(new Sheet()
                        {
                            Name = "Sheet",
                            SheetId = 1,
                            Id = document.WorkbookPart.GetIdOfPart(worksheetPart)
                        });

                        // this is for Sheets
                        writer1.WriteEndElement();
                        // this is for Workbook
                        writer1.WriteEndElement();
                        writer1.Close();
                    }
                }
                document.Close();
            }
        }

        private void WriteRow(List<string> columns, OpenXmlWriter writer, int rowNumber)
        {
            List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>();
            attributes.Add(new OpenXmlAttribute("r", null, rowNumber.ToString()));
            writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);
            foreach (string column in columns)
            {
                attributes = new List<OpenXmlAttribute>();
                attributes.Add(new OpenXmlAttribute("t", null, "str"));
                writer.WriteStartElement(new Cell(), attributes);
                writer.WriteElement(new CellValue(column));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private void WriteRows(List<Row> rows, OpenXmlWriter writer, int startRowNumber)
        {
            List<OpenXmlAttribute> attributes;
            int currentRowNumber = startRowNumber;
            foreach (var row in rows)
            {
                WriteRow(row.Columns, writer, currentRowNumber);
                currentRowNumber++;
            }
        }
    }
}
