using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EntityFramework.Utilities;

namespace TestingDbProject
{
    public class BulkInsertToDb
    {
        public void BulkInsert(IEnumerable<FakeObject> entities, int subBatchLength = 1, string warnMessagePrefix = "")
        {
            if (subBatchLength == 1)
            {
                BulkInsert(entities, warnMessagePrefix);
            }
            else
            {
                if (entities != null)
                {
                    var enumerable = entities as FakeObject[] ?? entities.ToArray();
                    var chunkSize = enumerable.Count() / subBatchLength + 1;
                    enumerable.Chunk(chunkSize).ToList().ForEach(x => BulkInsert(x, warnMessagePrefix));
                }
            }
        }

        protected void BulkInsert(IEnumerable<FakeObject> entities, string warnMessagePrefix = "")
        {
            FakeDbContext db = new FakeDbContext();
            if (entities == null)
                return;

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            try
            {
                EFBatchOperation.For(db, db.Objects).InsertAll(entities);

                stopWatch.Stop();
                Console.WriteLine(warnMessagePrefix + " : Bulk Insert duration in seconds: {0}", stopWatch.Elapsed.TotalSeconds);
            }
            catch (SqlException e)
            {
                stopWatch.Stop();
                Console.WriteLine(warnMessagePrefix + " : Bulk Insert duration in seconds: {0}", stopWatch.Elapsed.TotalSeconds);
                Console.WriteLine(e.Message, e);
                throw;
            }
        }
    }
}
