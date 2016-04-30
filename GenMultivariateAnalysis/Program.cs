using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using static System.Console;

namespace GenMultivariateAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 0) CreateTemplate(args[0], $"{Path.GetFileNameWithoutExtension(args[0])}.xlsx");
            else { WriteLine("csvファイルパスを入力");  var csv = ReadLine(); CreateTemplate(csv, $"{Path.GetFileNameWithoutExtension(csv)}.xlsx"); }
            WriteLine("Success! Generated a Excel Template!");
        }

        /// <summary>
        /// Read Multivariate Analysis of CSV Data 
        /// </summary>
        /// <param name="fileName">CSV File Path</param>
        /// <returns>If process is success, return Data else then return null</returns>
        static List<List<string>> ReadCSV(string fileName)
        {
            try
            {
                if (!File.Exists(fileName)) return null;
                var data = new List<List<string>>();
                using (var sr = new StreamReader(fileName))
                {
                    // Read to EOF
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine().Split(',');
                        data.Add(line.ToList());
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Error.WriteLine(ex);
                return null;
            }
        }

        static void CreateTemplate(string sourceFileName, string destFileName)
        {
            try
            {
                var data = ReadCSV(sourceFileName);
                if (data == null || data.Count == 0) return;

                // Create New WorkBook and WorkSheet
                var workBook = new XLWorkbook();
                var workSheet = workBook.Worksheets.Add("Sheet1");

                // Count of population, variables
                var population = data.Count;
                var variablesNum = data[0].Count;

                // Set Size
                for (int i = 0; i < variablesNum; i++)
                    workSheet.Column(i + 2).Width = 17;

                // Write Source Data and Calc Average and Variance
                {
                    int OffsetX = 2;
                    int OffsetY = 3;

                    // Set Labels
                    workSheet.Cell("A1").Value = "変量";
                    workSheet.Cell("B1").Value = "実測値";
                    for (var i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY - 2, i + OffsetX).Value = $"説明変量{i}";

                    workSheet.Cell("A2").Value = "個体";
                    workSheet.Cell("B2").Value = "y";
                    for (var i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY - 1, i + OffsetX).Value = $"x{i}";

                    // Write Source Data
                    for (var i = 0; i < population; i++)
                    {
                        workSheet.Cell(i + OffsetY, 1).Value = $"{(char)('A' + i)}";
                        for (var j = 0; j < variablesNum; j++)
                            workSheet.Cell(i + OffsetY, j + OffsetX).Value = data[i][j];
                    }

                    // Calc Average
                    workSheet.Cell(population + OffsetY, 1).Value = "平均";
                    for (var i = 0; i < variablesNum; i++)
                        workSheet.Cell(population + OffsetY, i + OffsetX).FormulaA1 = $"=AVERAGE({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})";

                    // Calc Variance
                    workSheet.Cell(population + OffsetY + 1, 1).Value = "分散";
                    for (var i = 0; i < variablesNum; i++)
                        workSheet.Cell(population + OffsetY + 1, i + OffsetX).FormulaA1 = $"=VAR({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})";
                }

                // Write Normalize Data and Calc Average and Variance
                {
                    int OffsetX = 2;
                    int OffsetY = population + 7;

                    // Set Labels
                    workSheet.Cell(OffsetY - 1, OffsetX).Value = "y";
                    for (var i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY - 1, i + OffsetX).Value = $"x{i}";

                    // Write Normalize Data
                    for (var i = 0; i < population; i++)
                    {
                        workSheet.Cell(i + OffsetY, 1).Value = $"{(char)('A' + i)}";
                        for (var j = 0; j < variablesNum; j++)
                            workSheet.Cell(i + OffsetY, j + OffsetX).FormulaA1 = $"=({(char)('B' + j)}{OffsetY - population - 4 + i}-{(char)('B' + j)}${OffsetY - 4})/SQRT({(char)('B' + j)}${OffsetY - 3})";
                    }

                    // Calc Average
                    workSheet.Cell(population + OffsetY, 1).Value = "平均";
                    for (var i = 0; i < variablesNum; i++)
                        workSheet.Cell(population + OffsetY, i + OffsetX).FormulaA1 = $"=AVERAGE({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})\n";

                    // Calc Variance
                    workSheet.Cell(population + OffsetY + 1, 1).Value = "分散";
                    for (var i = 0; i < variablesNum; i++)
                        workSheet.Cell(population + OffsetY + 1, i + OffsetX).FormulaA1 = $"=VAR({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})\n";
                }

                // Write Calc Variance/Covariance Matrix
                {
                    int OffsetX = 2;
                    int OffsetY = population * 2 + 11;
                    int NormX = OffsetX;
                    int NormY = OffsetY - population - 4;

                    // Set Labels
                    workSheet.Cell(OffsetY - 1, OffsetX).Value = "y";
                    for (var i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY - 1, i + OffsetX).Value = $"x{i}";

                    workSheet.Cell(OffsetY, OffsetX - 1).Value = "y";
                    for (var i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY + i, OffsetX - 1).Value = $"x{i}";

                    // Write Variace/Covariance
                    for (var i = 0; i < variablesNum; i++)
                    {
                        for (var j = 0; j < variablesNum; j++)
                        {
                            var arg1 = $"{(char)('B' + j)}{NormY}:{(char)('B' + j)}{NormY + population - 1}";
                            var arg2 = $"{(char)('B' + i)}{NormY}:{(char)('B' + i)}{NormY + population - 1}";
                            workSheet.Cell(OffsetY + i, OffsetX + j).FormulaA1 = $"=COVAR({arg1},{arg2})*{population}/({population}-1)";
                        }
                    }
                }

                // Write Calc Matrix
                {
                    int OffsetX = 1;
                    int OffsetY = population * 2 + variablesNum + 13;
                    int MatX = OffsetX;
                    int MatY = OffsetY - variablesNum - 2;

                    // Write Inverse
                    var arg = $"{(char)('B' + 1)}{MatY + 1}:{(char)('B' + variablesNum - 1)}{MatY + variablesNum - 1}";
                    workSheet.Range(OffsetY, OffsetX, OffsetY + variablesNum - 2, OffsetX + variablesNum - 2).FormulaA1 = $"=MINVERSE({arg})";

                    // Write value
                    for (int i = 1; i < variablesNum; i++)
                        workSheet.Cell(OffsetY + i - 1, OffsetX + variablesNum).FormulaA1 = $"B{MatY + i}";

                    // Write Calc result
                    var arg1 = $"A{OffsetY}:{(char)('A' + variablesNum - 2)}{OffsetY + variablesNum - 2}";
                    var arg2 = $"{(char)('B' + OffsetX + variablesNum - 2)}{OffsetY}:{(char)('B' + OffsetX + variablesNum - 2)}{OffsetY + variablesNum - 2}";
                    workSheet.Range(OffsetY, OffsetX + variablesNum + 2, OffsetY + variablesNum - 2, OffsetX + variablesNum + 2).FormulaA1 = $"=MMULT({arg1}, {arg2})";
                }

                // Write Calc Predicate/Error
                {
                    int OffsetX = 2;
                    int OffsetY = population * 2 + variablesNum * 2 + 14;
                    int MatX = OffsetX + variablesNum - 1;
                    int MatY = OffsetY - variablesNum - 1;
                    int NormX = 2;
                    int NormY = population + 7;

                    // Set Labels
                    workSheet.Cell($"B{OffsetY - 1}").Value = "Y";
                    workSheet.Cell($"C{OffsetY - 1}").Value = "E";

                    for (var i = 0; i < population; i++)
                    {
                        workSheet.Cell(i + OffsetY, 1).Value = $"{(char)('A' + i)}";
                        var y = "=";
                        for (var j = 0; j < variablesNum - 1; j++)
                            y += (j == 0 ? "" : "+") + $"{(char)('B' + j + 1)}{NormY + i}*{(char)('B' + MatX)}{MatY + j}";
                        workSheet.Cell(i + OffsetY, 2).FormulaA1 = y;
                        workSheet.Cell(i + OffsetY, 3).FormulaA1 = $"B{NormY + i}-B{OffsetY + i}";
                    }

                    // Calc Average
                    workSheet.Cell(population + OffsetY, 1).Value = "平均";
                    for (var i = 0; i < 2; i++)
                        workSheet.Cell(population + OffsetY, i + OffsetX).FormulaA1 = $"=AVERAGE({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})\n";

                    // Calc Variance
                    workSheet.Cell(population + OffsetY + 1, 1).Value = "分散";
                    for (var i = 0; i < 2; i++)
                        workSheet.Cell(population + OffsetY + 1, i + OffsetX).FormulaA1 = $"=VAR({(char)('B' + i)}${OffsetY}:{(char)('B' + i)}${population + OffsetY - 1})\n";
                }

                // Write Analysis of Variance Table
                {
                    int OffsetX = 1;
                    int OffsetY = population * 3 + variablesNum * 2 + 18;
                    int NormX = 2;
                    int NormY = population + 7;
                    int PredX = 2;
                    int PredY = population * 2 + variablesNum * 2 + 14;

                    // Set Labels
                    workSheet.Cell($"A{OffsetY - 1}").Value = "平方和";
                    workSheet.Cell($"B{OffsetY - 1}").Value = "自由度";
                    workSheet.Cell($"C{OffsetY - 1}").Value = "不偏分散";
                    workSheet.Cell($"D{OffsetY - 1}").Value = "分散比";

                    var arg = "";
                    for (int i = 0; i < population; i++)
                        arg += (i == 0 ? "" : ",") + $"B{PredY + i}-B{PredY + population}";
                    workSheet.Cell(OffsetY + 0, OffsetX + 0).FormulaA1 = $"=SUMSQ({arg})"; // Sr
                    workSheet.Cell(OffsetY + 1, OffsetX + 0).FormulaA1 = $"=SUMXMY2(B{NormY}:B{NormY + population - 1}, B{PredY}:B{PredY + population - 1})"; // Se
                    workSheet.Cell(OffsetY + 2, OffsetX + 0).FormulaA1 = $"=SUMSQ(B{NormY}:B{NormY + population - 1})"; // St
                    workSheet.Cell(OffsetY + 0, OffsetX + 1).Value = variablesNum - 1;              // p
                    workSheet.Cell(OffsetY + 1, OffsetX + 1).Value = population - variablesNum;     // n-p-1
                    workSheet.Cell(OffsetY + 2, OffsetX + 1).Value = population - 1;                // n-1
                    workSheet.Cell(OffsetY + 0, OffsetX + 2).FormulaA1 = $"A{OffsetY}/B{OffsetY}";
                    workSheet.Cell(OffsetY + 1, OffsetX + 2).FormulaA1 = $"A{OffsetY+1}/B{OffsetY+1}";
                    workSheet.Cell(OffsetY + 0, OffsetX + 3).FormulaA1 = $"C{OffsetY}/C{OffsetY+1}";
                }

                // Save file as destFileName and Close object
                workBook.SaveAs(destFileName);
                workBook.Dispose();
            }
            catch (Exception ex)
            {
                Error.WriteLine(ex);
                Environment.Exit(1);
            }
        }
    }
}
