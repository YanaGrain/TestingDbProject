using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace TestingDbProject
{
    public class ExcelGenerator
    {
        protected SpreadsheetDocument Document { get; set; }
        protected WorksheetPart WorksheetPart { get; set; }
        protected OpenXmlWriter Writer { get; set; }
        protected FileStream OutputStream { get; set; }
        protected Cell Cell { get; set; }

        protected string FilePath = Directory.GetCurrentDirectory() + @"\Excel";
        

        public ExcelGenerator()
        {
            this.OutputStream = new FileStream(this.FilePath, FileMode.OpenOrCreate);

            this.Document = SpreadsheetDocument.Create(this.OutputStream, SpreadsheetDocumentType.Workbook);
            this.Document.CompressionOption = CompressionOption.SuperFast;

            this.Document.AddWorkbookPart();
            this.WorksheetPart = this.Document.WorkbookPart.AddNewPart<WorksheetPart>();

            //var stylesPart = this.Document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            //var styles = new ReportExcelStyles();
            //styles.Save(stylesPart);

            this.Writer = OpenXmlWriter.Create(this.WorksheetPart);

        }

        public static OpenXmlAttribute GetNewRowAttribute(long rowIndex)
        {
            return new OpenXmlAttribute("r", null, rowIndex.ToString(CultureInfo.InvariantCulture));
        }

        public static void WriteNewRow(List<OpenXmlAttribute> attributes, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Row(), attributes);
        }

        public void WriteGeneralCell(OpenXmlWriter writer, string header, int rowIndex, string value, uint? styleIndex = null)
        {
            this.ResetCell();

            this.Cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "inlineStr"));
            this.Cell.CellReference = header + rowIndex;
            this.Cell.InlineString.Text.Text = value ?? string.Empty;

            if (styleIndex.HasValue)
            {
                this.Cell.StyleIndex = styleIndex.Value;
            }

            WriteCellSafe(this.Cell, writer);
        }

        private void ResetCell()
        {
            Cell.StyleIndex = null;
        }
    }
}
