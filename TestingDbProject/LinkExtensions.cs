using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingDbProject
{
   public static class LinkExtensions
    {
        public static IEnumerable<IEnumerable<ExampleObject>> Chunk<ExampleObject>(this IEnumerable<ExampleObject> source, int chunksize)
            {
                while (source.Any())
                {
                    yield return source.Take(chunksize);
                    source = source.Skip(chunksize);
                }
            }
    }
}
