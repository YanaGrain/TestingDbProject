using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using TestingDbProject;

namespace AppForTestingBigData
{
    class Program
    {
        private static Logger log = LogManager.GetCurrentClassLogger();

        private static readonly Random random = new Random();
        private static readonly object syncLock = new object();
        static void Main(string[] args)
        {
            int threadNumber = RandomNumber(30, 50);

            log.Info("Threads: " + threadNumber);
            Console.WriteLine("Threads: " + threadNumber);

            Parallel.For(0, 1, CreateExcelReport);
        }

        static void CreateExcelReport(int x)
        {
            log.Info("Thread " + x + " started");
            Console.WriteLine("Thread " + x + " started");

            string filePath = @"e:\SolbegSoft\Excel\" + Guid.NewGuid() + ".xlsx";
            int rowNumber = RandomNumber(300000, 500000);
            int skipNumber = RandomNumber(0, 1500000);

            log.Info(filePath + " " + skipNumber + " " + rowNumber);
            Console.WriteLine(filePath + " " + skipNumber + " " + rowNumber);

            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                xl.CompressionOption = CompressionOption.SuperFast;
                //IEnumerable<FakeObject> objects = db.Objects.OrderBy(ob => ob.Id).Skip(skipNumber).Take(rowNumber);
                xl.AddWorkbookPart();
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();



                using (var oxw = OpenXmlWriter.Create(wsp))
                {
                    oxw.WriteStartElement(new Worksheet());
                    oxw.WriteStartElement(new SheetData());

                    int chunkSize = 40000;
                    var ostatok = rowNumber%chunkSize;
                    int chunksCount = ostatok == 0 ? (rowNumber / chunkSize) : (rowNumber / chunkSize + 1);
                    List<string> headers = new List<string>();
                    headers.Add("1");
                    headers.Add("2");
                    headers.Add("3");
                    headers.Add("4");
                    headers.Add("5");
                    headers.Add("6");
                    headers.Add("7");
                    headers.Add("8");
                    headers.Add("9");
                    headers.Add("10");
                    headers.Add("11");
                    WriteRow(headers, oxw, 1);
                    try
                    {
                        Enumerable.Range(0, chunksCount).ToList().ForEach(number =>
                        {
                            var currentChunkSize = number == chunksCount - 1 ? ostatok : chunkSize;
                            FakeDbContext db = new FakeDbContext();

                            log.Info("Thread " + x + " chunk " + number);
                            Console.WriteLine("Thread " + x + " chunk " + number);

                            var chunk =
                                db.Objects.OrderBy(ob => ob.Id)
                                    .Skip(skipNumber + number*chunkSize)
                                    .Take(currentChunkSize)
                                    .ToList();
                            WriteRow(chunk, oxw, number*chunkSize+2);
                        });
                    }
                    catch (Exception e)
                    {
                        log.Error(e.Message);
                    }

                    // this is for SheetData
                    oxw.WriteEndElement();
                    // this is for Worksheet
                    oxw.WriteEndElement();
                    oxw.Close();

                    using (var oxw1 = OpenXmlWriter.Create(xl.WorkbookPart))
                    {
                        oxw1.WriteStartElement(new Workbook());
                        oxw1.WriteStartElement(new Sheets());

                        oxw1.WriteElement(new Sheet()
                        {
                            Name = "Sheet1",
                            SheetId = 1,
                            Id = xl.WorkbookPart.GetIdOfPart(wsp)
                        });

                        // this is for Sheets
                        oxw1.WriteEndElement();
                        // this is for Workbook
                        oxw1.WriteEndElement();
                        oxw1.Close();
                    }
                }
                xl.Close();
            }

            log.Info("Thread " + x + " finished");
            Console.WriteLine("Thread " + x + " finished");
        }
        public static int RandomNumber(int min, int max)
        {
            lock (syncLock)
            { // synchronize
                return random.Next(min, max);
            }
        }

        static void WriteRow(IEnumerable<FakeObject> objects, OpenXmlWriter oxw, int i)
        {
            List<OpenXmlAttribute> oxa;
            foreach (var fakeObject in objects)
            {
                //FakeObject fakeObject = objects[i - 1];

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

                    oxa.Add(new OpenXmlAttribute("t", null, "str"));

                    oxw.WriteStartElement(new Cell(), oxa);

                    oxw.WriteElement(new CellValue(Convert.ToString(propValues[j])));

                    // this is for Cell
                    oxw.WriteEndElement();
                }

                // this is for Row
                oxw.WriteEndElement();
                i++;
            }
        }

        static void WriteRow(List<string> columns, OpenXmlWriter writer, int rowNumber)
        {
            List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>();
            attributes.Add(new OpenXmlAttribute("r", null, rowNumber.ToString()));
            writer.WriteStartElement(new Row(), attributes);
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
    }
}

