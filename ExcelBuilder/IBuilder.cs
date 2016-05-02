using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilder
{
    public interface IBuilder
    {
        Container Build();
    }
}
