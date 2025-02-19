using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plagiarism_Validation
{
    internal class DisjointSet
    {

        private Dictionary<string, string> parent;
        private Dictionary<string, int> rank;

        public DisjointSet(IEnumerable<string> elements)
        {
            parent = new Dictionary<string, string>();
            rank = new Dictionary<string, int>();
            foreach (var element in elements)
            {
                parent[element] = element;
                rank[element] = 0;
            }
        }

        public string Find(string x)
        {
            if (parent[x] != x)
            {
                parent[x] = Find(parent[x]);
            }
            return parent[x];
        }

        public void Union(string x, string y)
        {
            var rootX = Find(x);
            var rootY = Find(y);

            if (rootX != rootY)
            {
                if (rank[rootX] > rank[rootY])
                {
                    parent[rootY] = rootX;
                }
                else if (rank[rootX] < rank[rootY])
                {
                    parent[rootX] = rootY;
                }
                else
                {
                    parent[rootY] = rootX;
                    rank[rootX]++;
                }
            }
        }

    }
}
