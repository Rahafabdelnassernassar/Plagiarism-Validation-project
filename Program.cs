using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;
using System.Drawing.Imaging;
//using static System.Runtime.InteropServices.JavaScript.JSType;
//using System.ComponentModel;
//using DocumentFormat.OpenXml.Vml; 


namespace Plagiarism_Validation
{
    internal class Program
    {

        //////////////////////////////        Reading excel file and constructing graph          ///////////////////////////////

        // node -> neighbour, semilarity, #lines, nodeSemilarity
        public static Dictionary<string, List<Tuple<string, int, int, int>>> graph;
        // hyperlinks
        public static Dictionary<string, string> URLs;
        public static void constructGraph(List<pairsOfNodes> matchingPairs)
        {

            graph = new Dictionary<string, List<Tuple<string, int, int, int>>>();

            foreach (pairsOfNodes pair in matchingPairs)
            {
                if (!graph.ContainsKey(pair.File1Path))
                {
                    graph[pair.File1Path] = new List<Tuple<string, int, int, int>>();
                }
                if (!graph.ContainsKey(pair.File2Path))
                {
                    graph[pair.File2Path] = new List<Tuple<string, int, int, int>>();
                }

                Tuple<string, int, int, int> neighbor2 = new Tuple<string, int, int, int>(pair.File2Path, pair.Similarity2, pair.MatchedLines, pair.Similarity1);
                graph[pair.File1Path].Add(neighbor2);

                Tuple<string, int, int, int> neighbor1 = new Tuple<string, int, int, int>(pair.File1Path, pair.Similarity1, pair.MatchedLines, pair.Similarity2);
                graph[pair.File2Path].Add(neighbor1);
            }
        }
        static List<pairsOfNodes> ReadMatchingPairsFromFile(string filePath)
        {
            List<pairsOfNodes> matchingPairs = new List<pairsOfNodes>();

            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(filePath);
            Worksheet worksheet = document.Workbook.Worksheets[0];

            for (int i = 1; i <= worksheet.UsedRangeRowMax; i++)
            {
                Cell cell1 = worksheet.Cell(i, 0);
                string cellValue1 = cell1.Value.ToString();
                string[] parts1 = cellValue1.Split('(', ')', '%');

                string filePath1 = parts1[0];
                string hyperlink1 = parts1[0]; // Assuming the hyperlink is the same as the file path

                int similarity1 = Int32.Parse(parts1[1]);

                Cell cell2 = worksheet.Cell(i, 1);
                string cellValue2 = cell2.Value.ToString();
                string[] parts2 = cellValue2.Split('(', ')', '%');

                string filePath2 = parts2[0];
                string hyperlink2 = parts2[0]; // Assuming the hyperlink is the same as the file path

                int similarity2 = Int32.Parse(parts2[1]);

                Cell cell3 = worksheet.Cell(i, 2);
                string cellValue3 = cell3.Value.ToString();
                int matchedLines = Int32.Parse(cellValue3);

                pairsOfNodes pair = new pairsOfNodes(filePath1, hyperlink1, filePath2, hyperlink2, similarity1, similarity2, matchedLines);
                matchingPairs.Add(pair);
            }
            constructGraph(matchingPairs);
            ReadHyperlinksFromExcel(filePath);

            document.Close();

            return matchingPairs;
        }

        ///////////////////////////////////      saving hyperlinks      /////////////////////////////////////////
        private static void ReadHyperlinksFromExcel(string filePath)
        {
            URLs = new Dictionary<string, string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell.Hyperlink != null)
                        {
                            var cellValue = cell.Text;
                            var url = cell.Hyperlink.AbsoluteUri;
                            URLs[cellValue] = url;
                        }
                    }
                }
            }
            //PrintURLs();
        }

        ///////////////////////////////////////           groups and statistics         ///////////////////////////////////////////
        public static Dictionary<List<string>, (float, int)> findingGroups(Dictionary<string, List<Tuple<string, int, int, int>>> graph, ref List<Dictionary<string, List<Tuple<string, int, int, int>>>> groupsGraph, ref List<float> avgSimilarity)
        {
            //HashSet<string> visited = new HashSet<string>();
            // List<List<string>> groups = new List<List<string>>();
            //  Dictionary<List<string>, float> groupSimilarity = new Dictionary<List<string>, float>();
            Dictionary<string, string> status = graph.Keys.ToDictionary(key => key, value => "white");
            Dictionary<List<string>, (float, int)> groupS = new Dictionary<List<string>, (float, int)>();

            float avarage;

            foreach (var node in graph.Keys)
            {
                if (status[node] == "white")
                {
                    List<string> group = new List<string>();
                    Dictionary<string, List<Tuple<string, int, int, int>>> groupGraph = new Dictionary<string, List<Tuple<string, int, int, int>>>();
                    float totalScore = 0;
                    int totalEdges = 0, groupCount = 1;
                    DFS(node, graph, ref group, status, ref totalScore, ref totalEdges, ref groupCount, ref groupGraph);
                    //groups.Add(group);

                    avarage = totalScore / totalEdges;
                    float roundedNumber = (float)Math.Round(avarage, 1);
                    // groupSimilarity[group] = roundedNumber;
                    groupS[group] = (roundedNumber, groupCount);
                    groupsGraph.Add(groupGraph);
                    avgSimilarity.Add(totalScore / totalEdges);
                }
            }

            Dictionary<List<string>, (float, int)> sortedGroup = SortGroupsByAvgSim(groupS);
            var sortedGroupS = SortNodesAscending(sortedGroup);
            // statFile(sortedGroupS);
            // PrintGroups(sortedGroupS);
            return sortedGroupS;
        }

        private static void DFS(string node, Dictionary<string, List<Tuple<string, int, int, int>>> graph, ref List<string> group, Dictionary<string, string> status, ref float totalScore, ref int totalEdges, ref int groupCount, ref Dictionary<string, List<Tuple<string, int, int, int>>> groupGraph)
        {
            status[node] = "gray";
            foreach (var neighbour in graph[node])
            {

                if (status[neighbour.Item1] == "white")
                {
                    DFS(neighbour.Item1, graph, ref group, status, ref totalScore, ref totalEdges, ref groupCount, ref groupGraph);
                    groupCount++;

                    if (!groupGraph.ContainsKey(node))
                    {
                        groupGraph[node] = new List<Tuple<string, int, int, int>>();
                    }
                    if (!groupGraph.ContainsKey(neighbour.Item1))
                    {
                        groupGraph[neighbour.Item1] = new List<Tuple<string, int, int, int>>();
                    }
                    groupGraph[node].Add(Tuple.Create(neighbour.Item1, neighbour.Item2, neighbour.Item3, neighbour.Item4));
                    groupGraph[neighbour.Item1].Add(Tuple.Create(node, neighbour.Item4, neighbour.Item3, neighbour.Item2));

                }
                if (status[neighbour.Item1] == "gray" || status[neighbour.Item1] == "black")
                {
                    totalScore += neighbour.Item2;
                    totalEdges++;

                    if (!groupGraph.ContainsKey(node))
                    {
                        groupGraph[node] = new List<Tuple<string, int, int, int>>();
                    }
                    if (!groupGraph.ContainsKey(neighbour.Item1))
                    {
                        groupGraph[neighbour.Item1] = new List<Tuple<string, int, int, int>>();
                    }
                    groupGraph[node].Add(Tuple.Create(neighbour.Item1, neighbour.Item2, neighbour.Item3, neighbour.Item4));
                    groupGraph[neighbour.Item1].Add(Tuple.Create(node, neighbour.Item4, neighbour.Item3, neighbour.Item2));
                }
            }
            status[node] = "black";

            int num = GetNumericPart(node);
            group.Add(num.ToString());
        }

        private static int GetNumericPart(string path)
        {
            string numericPart = Regex.Match(path, @"\d+").Value;

            if (!string.IsNullOrEmpty(numericPart))
            {
                int numericValue = int.Parse(numericPart);
                return numericValue;
            }

            return -1;
        }

        private static Dictionary<List<string>, (float, int)> SortGroupsByAvgSim(Dictionary<List<string>, (float, int)> groupS)
        {
            var sortedDictionary = groupS.OrderByDescending(pair => pair.Value.Item1).ToDictionary(pair => pair.Key, pair => pair.Value);

            return sortedDictionary;
        }

        static Dictionary<List<string>, (float, int)> SortNodesAscending(Dictionary<List<string>, (float, int)> groupS)
        {
            var sortedGroupS = groupS.ToDictionary(
                item => item.Key.OrderBy(str => int.Parse(str)).ToList(),
                item => item.Value
            );
            //PrintGroups(sortedGroupS);
            return sortedGroupS;
        }

        private static List<Dictionary<string, List<Tuple<string, int, int, int>>>> SortMaster_Slave(List<float> avgSimilarity, List<Dictionary<string, List<Tuple<string, int, int, int>>>> groupsGraph)
        {
            var combinedList = groupsGraph.Zip(avgSimilarity, (graph, similarity) => new { Graph = graph, Similarity = similarity });

            // Sort in descending order based on avgSimilarity
            var sortedList = combinedList.OrderByDescending(item => item.Similarity).ToList();

            // Extract the sorted groupsGraph entries
            var sortedGroupsGraph = sortedList.Select(item => item.Graph).ToList();

            return sortedGroupsGraph;
        }

        /////////////////////////////////////////////           MST       /////////////////////////////////////////////

        public static List<Tuple<string, string, int, int, int>> FindMST(List<Dictionary<string, List<Tuple<string, int, int, int>>>> groupsGraph)
        {

            var MST = new List<Tuple<string, string, int, int, int>>();

            foreach (var groupGraph in groupsGraph)
            {
                // node1 , node2 , similarity2 , similarity1 , lineMaches
                var edges = new List<Tuple<string, string, int, int, int>>();
                var seenEdges = new HashSet<Tuple<string, string, int, int, int>>();

                foreach (var node in groupGraph.Keys)
                {
                    foreach (var neighbor in groupGraph[node])
                    {
                        var edge = Tuple.Create(node, neighbor.Item1, neighbor.Item2, neighbor.Item4, neighbor.Item3);
                        var reverseEdge = Tuple.Create(neighbor.Item1, node, neighbor.Item2, neighbor.Item4, neighbor.Item3);

                        if (!seenEdges.Contains(reverseEdge))
                        {
                            edges.Add(edge);
                            seenEdges.Add(edge);
                        }
                    }
                }

                edges = edges.OrderByDescending(e => e.Item4).ThenByDescending(e => e.Item5).ToList();

                var disjointSet = new DisjointSet(groupGraph.Keys);

                var groupMST = new List<Tuple<string, string, int, int, int>>();

                foreach (var edge in edges)
                {
                    var (node1, node2, similarity1, similarity2, matchedLines) = edge;

                    if (disjointSet.Find(node1) != disjointSet.Find(node2))
                    {
                        groupMST.Add(edge);
                        disjointSet.Union(node1, node2);
                    }
                }
                groupMST.Sort((t1, t2) => t2.Item5.CompareTo(t1.Item5));

                foreach (var edge in groupMST)
                {
                    MST.Add(edge);
                }
            }

            return MST;
        }

        ////////////////////////////////////////        Generating stats & mst files      //////////////////////////////////////

        public static void statFile(Dictionary<List<string>, (float, int)> groupS)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string folderPath = "D:\\learning\\Kolia\\3rd Year\\2nd semester\\Algo\\project\\Our Result Files";
            string fileName = "statFile.xlsx";
            string filePath = Path.Combine(folderPath, fileName);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Components");

                // Column headers
                worksheet.Cells[1, 1].Value = "Component Index";
                worksheet.Cells[1, 2].Value = "Vertices";
                worksheet.Cells[1, 3].Value = "Average Similarity";
                worksheet.Cells[1, 4].Value = "Component Count";

                // Fill data
                int rowIndex = 2;
                int componentIndex = 1;

                foreach (var item in groupS)
                {

                    worksheet.Cells[rowIndex, 1].Value = componentIndex;

                    worksheet.Cells[rowIndex, 2].Value = string.Join(", ", item.Key);

                    worksheet.Cells[rowIndex, 3].Value = item.Value.Item1;

                    worksheet.Cells[rowIndex, 4].Value = item.Value.Item2;

                    //worksheet.Column(1).AutoFit();
                    //worksheet.Column(2).AutoFit();
                    //worksheet.Column(3).AutoFit();
                    //worksheet.Column(4).AutoFit();

                    rowIndex++;
                    componentIndex++;
                }
                    worksheet.Column(2).AutoFit();

                // Save the workbook to the specified folder and file path
                package.SaveAs(new FileInfo(filePath));

                //Console.WriteLine("Excel file saved successfully.");
            }

            Console.WriteLine("stat file created and filled with data successfully.");
        }

        public static void MST_file(List<Tuple<string, string, int, int, int>> mst)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                var worksheet = excelPackage.Workbook.Worksheets.Add("mst");

                // Column headers
                worksheet.Cells[1, 1].Value = "File 1";
                worksheet.Cells[1, 2].Value = "File 2";
                worksheet.Cells[1, 3].Value = "Line Matches";
                int rowIndex = 2;

                foreach (var edge in mst)
                {

                    string formattedValue1 = $"{edge.Item1}({edge.Item4}%)";
                    string formattedValue2 = $"{edge.Item2}({edge.Item3}%)";

                    worksheet.Cells[rowIndex, 1].Value = formattedValue1;
                    worksheet.Cells[rowIndex, 2].Value = formattedValue2;
                    worksheet.Cells[rowIndex, 3].Value = edge.Item5;

                    // Set hyperlinks 
                    try
                    { 
                    worksheet.Cells[rowIndex, 1].Hyperlink = new ExcelHyperLink(new Uri(URLs[formattedValue1]).ToString());
                    worksheet.Cells[rowIndex, 1].Style.Font.UnderLine = true;
                    worksheet.Cells[rowIndex, 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    } 
                    catch 
                    {
                        worksheet.Cells[rowIndex, 1].Hyperlink = new ExcelHyperLink(new Uri(formattedValue1).ToString());
                        worksheet.Cells[rowIndex, 1].Style.Font.UnderLine = true;
                        worksheet.Cells[rowIndex, 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }
                    try 
                    {
                        worksheet.Cells[rowIndex, 2].Hyperlink = new ExcelHyperLink(new Uri(URLs[formattedValue2]).ToString());
                        worksheet.Cells[rowIndex, 2].Style.Font.UnderLine = true;
                        worksheet.Cells[rowIndex, 2].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }
                    catch
                    {
                        worksheet.Cells[rowIndex, 2].Hyperlink = new ExcelHyperLink(new Uri(formattedValue2).ToString());
                        worksheet.Cells[rowIndex, 2].Style.Font.UnderLine = true;
                        worksheet.Cells[rowIndex, 2].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }

                    //worksheet.Column(1).AutoFit();
                    //worksheet.Column(2).AutoFit();
                    //worksheet.Column(3).AutoFit();

                    rowIndex++;
                }

                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();


                string folderPath = "D:\\learning\\Kolia\\3rd Year\\2nd semester\\Algo\\project\\Our Result Files";
                string fileName = "mstFile.xlsx";
                string filePath = Path.Combine(folderPath, fileName);
                excelPackage.SaveAs(new FileInfo(filePath));

                Console.WriteLine("mst file created and filled with data successfully.");
            }
        }

        ///////////////////////////////////////       Printing graph and Groups       /////////////////////////////////////////
        public static void PrintGraph(Dictionary<string, List<Tuple<string, int, int, int>>> graph)
        {
            foreach (var vertex in graph)
            {
                Console.WriteLine("Vertex: " + vertex.Key);
                Console.WriteLine("Adjacent vertices:");

                foreach (var neighbor in vertex.Value)
                {
                    Console.WriteLine("  - " + neighbor.Item1 + ", Weight: " + neighbor.Item2 + ", matchedLines: " + neighbor.Item3);
                }

                Console.WriteLine();
            }
        }

        public static void PrintGroups(Dictionary<List<string>, (float, int)> groupS)
        {
            Console.WriteLine("Groups:");
            foreach (var group in groupS.Keys)
            {
                Console.WriteLine("Group:");
                foreach (var vertex in group)
                {
                    Console.WriteLine(vertex);
                }
                Console.WriteLine();
                Console.WriteLine("with Similarity: " + groupS[group].Item1.ToString("F1"));
                Console.WriteLine("with Count: " + groupS[group].Item2);
                Console.WriteLine();
            }
        }

        static void Main(string[] args)
        {
            string filePath = "D:\\learning\\Kolia\\3rd Year\\2nd semester\\Algo\\project\\[3] Plagiarism Validation\\Test Cases\\Complete\\Easy\\2-Input.xlsx";
            // Total Time stopwatch
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            // Reading and constructing graph
            ReadMatchingPairsFromFile(filePath);

            // groups and statistics
            List<Dictionary<string, List<Tuple<string, int, int, int>>>> groupsGraph = new List<Dictionary<string, List<Tuple<string, int, int, int>>>>();
            List<float> avgSimilarity = new List<float>();
            Dictionary<List<string>, (float, int)> groups = findingGroups(graph, ref groupsGraph, ref avgSimilarity);

            // MST
            var sortedGroupsGraph = SortMaster_Slave(avgSimilarity, groupsGraph);

            // mst Algorithm & file generation stopwatch
            Stopwatch mst_algo_and_file_stopWatch = new Stopwatch();
            mst_algo_and_file_stopWatch.Start();

            var mst = FindMST(sortedGroupsGraph);

            // mst file stopwatch
            Stopwatch mst_file_stopWatch = new Stopwatch();
            mst_file_stopWatch.Start();

            // mst file generation and save
            MST_file(mst);

            mst_file_stopWatch.Stop();
            Console.WriteLine($"mst file time: {mst_file_stopWatch.ElapsedMilliseconds} ms");

            mst_algo_and_file_stopWatch.Stop();
            Console.WriteLine($"mst Algorithm & file time: {mst_algo_and_file_stopWatch.ElapsedMilliseconds} ms");

            // stat file stopwatch
            Stopwatch stat_file_stopWatch = new Stopwatch();
            stat_file_stopWatch.Start();

            // stat file generation and save
            statFile(groups);

            stat_file_stopWatch.Stop();
            Console.WriteLine($"stat file time: {stat_file_stopWatch.ElapsedMilliseconds} ms");

            stopwatch.Stop();
            Console.WriteLine($"Total Execution time: {stopwatch.ElapsedMilliseconds} ms");

        }
    }
}