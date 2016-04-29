using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestingDbProject;
using System.IO.Packaging;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Tests
{
   [TestClass]
    public class ExcelDocBuilder
    {
        [TestMethod]
        public void TestMethod()
        {
            int rowNumber = 10000;
            //Make a copy of the template file.
            string from = @"e:\SolbegSoft\Excel\template.xlsx";
            string to = @"e:\SolbegSoft\Excel\" + Guid.NewGuid() + ".xlsx";
            File.Copy(from, to, true);
            //Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(to, true))
            {
                myWorkbook.CompressionOption = CompressionOption.SuperFast;
                //Access the main Workbook part, which contains all references.
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;
                //Get the first worksheet. 
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                string origninalSheetId = workbookPart.GetIdOfPart(worksheetPart);
                WorksheetPart replacementPart = workbookPart.AddNewPart<WorksheetPart>();
                string replacementPartId = workbookPart.GetIdOfPart(replacementPart);
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart);

                // The SheetData object will contain all the data.
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                using (FakeDbContext db = new FakeDbContext())
                {
                    List<FakeObject> objects = db.Objects.Take(rowNumber).ToList();
                    //int numRows = objects.Count();
                    int numRows = rowNumber;
                    for (int rowIndex = 2; rowIndex < numRows; rowIndex++)
                    {
                        var props = typeof(FakeObject).GetProperties();
                        List<string> propValues = new List<string>();

                        FakeObject fakeObject = objects[rowIndex];

                        foreach (var prop in props)
                        {
                            //generator.WriteGeneralCell(generator.Writer, prop.Name, rowIndex, prop.ToString());
                            string value = prop.GetValue(fakeObject, null).ToString();
                            propValues.Add(value);
                        }

                        Row contentRow = CreateContentRow(rowIndex, propValues, reader, writer, numRows);

                        
                        

                        //Append new row to sheet data.
                        //sheetData.AppendChild(contentRow);
                    }
                }

                reader.Close();
                writer.Close();

                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Id.Value.Equals(origninalSheetId)).First();
                sheet.Id.Value = replacementPartId;
                workbookPart.DeletePart(worksheetPart);

                //worksheetPart.Worksheet.Save();
                myWorkbook.Close();
            }
        }

        string[] headerColumns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
        private Row CreateContentRow(long rowIndex, List<string> props, OpenXmlReader reader, OpenXmlWriter writer, int numRows)
        {
            //Create the new row.
            Row r = new Row();
            r.RowIndex = (UInt32)rowIndex;
            //First cell is a text cell, so create it and append it.
            Cell firstCell = CreateTextCell(headerColumns[0], props[0], rowIndex);
            r.AppendChild(firstCell);
            //Create the cells that contain the data.
            for (int i = 1; i < headerColumns.Length; i++)
            {
                Cell c = new Cell();
                c.CellReference = headerColumns[i] + rowIndex;
                CellValue v = new CellValue();
                v.Text = props[i];
                c.DataType = new EnumValue<CellValues>(CellValues.String);
                c.AppendChild(v);
                r.AppendChild(c);

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(SheetData))
                    {
                        if (reader.IsEndElement)
                            continue;
                        writer.WriteStartElement(new SheetData());

                        for (int row = 0; row < numRows; row++)
                        {
                            writer.WriteStartElement(r);
                            for (int col = 0; col < 11; col++)
                            {
                                writer.WriteElement(c);
                            }
                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement();
                    }
                    else
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                }


            }

            return r;
        }

        private Cell CreateTextCell(string header, string text, long index)
        {
            //Create a new inline string cell.
            Cell c = new Cell();
            c.DataType = CellValues.InlineString;
            c.CellReference = header + index;
            //Add text to the text cell.
            InlineString inlineString = new InlineString();
            Text t = new Text();
            t.Text = text;
            inlineString.AppendChild(t);
            c.AppendChild(inlineString);



            return c;
        }

    }
}
