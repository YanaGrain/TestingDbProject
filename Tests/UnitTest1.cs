using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestingDbProject;

namespace Tests
{
    [TestClass]
    public class UnitTest1
    {
        private string filePath = @"e:\SolbegSoft\Excel\" + Guid.NewGuid() + ".xlsx";
        [TestMethod]
        public void TestMethod1()
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                xl.CompressionOption = CompressionOption.SuperFast;
                int rowNumber = 80000;
                FakeDbContext db = new FakeDbContext();
                List<FakeObject> objects = db.Objects.Take(rowNumber).ToList();
                List<OpenXmlAttribute> oxa;
                OpenXmlWriter oxw;

                xl.AddWorkbookPart();
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

                oxw = OpenXmlWriter.Create(wsp);
                oxw.WriteStartElement(new Worksheet());
                oxw.WriteStartElement(new SheetData());

                for (int i = 1; i <= rowNumber; ++i)
                {
                    FakeObject fakeObject = objects[i-1];

                    oxa = new List<OpenXmlAttribute>();
                    // this is the row index
                    oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));

                    oxw.WriteStartElement(new Row(), oxa);

                    for (int j = 0; j <= 10; ++j)
                    {
                        var props = typeof(FakeObject).GetProperties();
                        List<string> propValues = new List<string>();
                        foreach (var prop in props)
                        {
                            //generator.WriteGeneralCell(generator.Writer, prop.Name, rowIndex, prop.ToString());
                            string value = prop.GetValue(fakeObject, null).ToString();
                            propValues.Add(value);
                        }

                        oxa = new List<OpenXmlAttribute>();
                        // this is the data type ("t"), with CellValues.String ("str")
                        oxa.Add(new OpenXmlAttribute("t", null, "str"));

                        // it's suggested you also have the cell reference, but
                        // you'll have to calculate the correct cell reference yourself.
                        // Here's an example:
                        //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                        oxw.WriteStartElement(new Cell(), oxa);

                        oxw.WriteElement(new CellValue(propValues[j].ToString()));

                        // this is for Cell
                        oxw.WriteEndElement();
                    }

                    // this is for Row
                    oxw.WriteEndElement();
                }

                // this is for SheetData
                oxw.WriteEndElement();
                // this is for Worksheet
                oxw.WriteEndElement();
                oxw.Close();

                oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.
                oxw.WriteElement(new Sheet()
                {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = xl.WorkbookPart.GetIdOfPart(wsp)
                });

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                xl.Close();
            }
        }
    }
}
