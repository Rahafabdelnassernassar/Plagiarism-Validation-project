using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plagiarism_Validation
{
    internal class pairsOfNodes
    {
        public string File1Path { get; set; }
        public string Hyperlink1 { get; set; }
        public string File2Path { get; set; }
        public string Hyperlink2 { get; set; }
        public int Similarity1 { get; set; }
        public int Similarity2 { get; set; }
        public int MatchedLines { get; set; }

        public pairsOfNodes(string file1Path, string hyperlink1, string file2Path, string hyperlink2, int similarity1, int similarity2, int matchedLines)
        {
            File1Path = file1Path;
            Hyperlink1 = hyperlink1;
            File2Path = file2Path;
            Hyperlink2 = hyperlink2;
            Similarity1 = similarity1;
            Similarity2 = similarity2;
            MatchedLines = matchedLines;
        }
    }

}