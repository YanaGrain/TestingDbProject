using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingDbProject
{
   public static class LinkExtensions
    {
        public static IEnumerable<IEnumerable<FakeObject>> Chunk<FakeObject>(this IEnumerable<FakeObject> source, int chunksize)
            {
                while (source.Any())
                {
                    yield return source.Take(chunksize);
                    source = source.Skip(chunksize);
                }
            }

        public static IEnumerable<IEnumerable<FakeObject>> Chunk<FakeObject>(this IEnumerable<FakeObject> source, int chunksize, int chunkNumber)
        {
            for (int i = 0; i < chunkNumber; i++)
            {
                yield return source.Take(chunksize);
                source = source.Skip(chunksize);
            }
        }
    }
}
