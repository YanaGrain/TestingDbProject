using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ExcelBuilder;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestingDbProject;

namespace Tests
{
    [TestClass]
    public class DbInitializer
    {
        [TestMethod]
        public void TestMethod()
        {
            Random rnd = new Random();
            var db = new BulkInsertToDb();
            //FakeObjects.ToList().ForEach(l => db.BulkInsert(l));

            foreach (var l in ExampleObjects)
            {
                db.BulkInsert(l);
            }
        }

        private IEnumerable<IEnumerable<ExampleObject>> ExampleObjects
        {
            get
            {
                Random rnd = new Random();

                for (int i = 0; i < 200; i++)
                {
                    List<ExampleObject> objects = new List<ExampleObject>();

                    for (int j = 0; j < 10000; j++)
                    {
                        objects.Add(BuildObject(rnd));
                    }

                    yield return objects;
                }
            }
        }

        private ExampleObject BuildObject(Random rnd)
        {
            return new ExampleObject()
            {
                Age = rnd.Next(100),
                City = Guid.NewGuid().ToString(),
                Name = Guid.NewGuid().ToString(),
                OnVacation = true,
                Prop1 = rnd.Next(1000),
                Prop2 = rnd.NextDouble(),
                Prop3 = rnd.Next(200000),
                Prop4 = rnd.Next(1500),
                Prop5 = Guid.NewGuid().ToString(),
                Prop6 = Guid.NewGuid().ToString()
            };
        }
    }
}
