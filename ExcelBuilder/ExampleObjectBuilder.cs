using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilder
{
    public class ExampleObjectBuilder : IBuilder
    {
        public const int RowNumber = 100; // get 100 rows from db

        public Container Build()
        {
            Container container = new Container();

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

            container.Headers = headers;

            List<ExampleObject> objects;
            using (ExampleDbContext db = new ExampleDbContext())
            {
               objects = db.Objects.Take(RowNumber).ToList();
            }
            
            List<Row> rows = new List<Row>();
            var props = typeof(ExampleObject).GetProperties();
            Row row;
            foreach (var obj in objects)
            {
                row = new Row();
                row.Columns = new List<string>();
                foreach (var prop in props)
                {
                    string value = prop.GetValue(obj, null).ToString();
                    row.Columns.Add(value);
                }
                rows.Add(row);
            }

            container.Rows = rows;

            return container;
        }
    }
}
