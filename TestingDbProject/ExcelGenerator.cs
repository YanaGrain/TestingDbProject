using System;
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
        public string[] headerColumns;

        public ExcelGenerator()
        {
            headerColumns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
        }

        public Row CreateContentRow(long rowIndex, List<string> properties)
        {
            //Create the new row.
            Row row = new Row();
            row.RowIndex = (UInt32)rowIndex;
            //First cell is a text cell, so create it and append it.
            Cell firstCell = CreateTextCell(headerColumns[0], properties[0], rowIndex);
            row.AppendChild(firstCell);
            //Create the cells that contain the data.
            for (int i = 1; i < headerColumns.Length; i++)
            {
                Cell c = new Cell();
                c.CellReference = headerColumns[i] + rowIndex;
                CellValue v = new CellValue();
                v.Text = properties[i];
                c.DataType = new EnumValue<CellValues>(CellValues.String);
                c.AppendChild(v);
                row.AppendChild(c);
            }
            return row;
        }

        private Cell CreateTextCell(string header, string text, long index)
        {
            //Create a new inline string cell.
            Cell c = new Cell();
            c.DataType = CellValues.String;
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
