using System;
using System.Text;
using System.Collections.Generic;
using ExcelBuilder;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
   [TestClass]
    public class ExcelBuilderTest
    {
        [TestMethod]
        public void ShouldGenerateReport()
        {
            ExampleObjectBuilder builder = new ExampleObjectBuilder();
            ReportGenerator generator = new ReportGenerator(builder);
            generator.Generate();
        }
    }
}
