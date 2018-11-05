namespace Excellent
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Text;

    using ClosedXML.Excel;

    using KellermanSoftware.CompareNetObjects;

    using Serilog;

    using SmartFormat;

    public static class Utils
    {
        public static int Transform(string input, string output, string outputFormat)
        {
            Log.Information($"Processing '{input}'");

            var dataSet = input.GetData();
            var sheetsCount = dataSet?.Tables?.Count;
            Log.Information($"Found {sheetsCount} sheets\n");

            for (var i = 0; i < sheetsCount; i++)
            {
                var table = dataSet?.Tables[i];
                Log.Information($"Processing '{table.TableName}' sheet");
                var rows = table.GetRows<ExpandoObject>().ToList();
                if (rows?.Count > 0)
                {
                    Log.Information($"Row Count: {rows.Count}");
                    CheckDuplicates(rows);
                    var result = new StringBuilder();
                    result.AppendLine($"-- {table.TableName}");
                    foreach (var row in rows)
                    {
                        var val = Smart.Format(outputFormat, row);
                        result.AppendLine(val);
                    }

                    result.AppendLine();
                    Log.Information($"Writing output to '{output}'\n");
                    if (i == 0)
                    {
                        File.WriteAllText(output, result.ToString());
                    }
                    else
                    {
                        File.AppendAllText(output, result.ToString());
                    }
                }
            }

            return 0;
        }

        public static int Merge(IEnumerable<string> inputs, string output, bool keepRight, bool keepLeft)
        {
            Log.Information($"Processing '{string.Join(", ", inputs)}'");
            var tableDict = new ConcurrentDictionary<string, ConcurrentDictionary<string, ExpandoObject>>();
            foreach (var input in inputs)
            {
                var dataSet = input.GetData();
                var sheetsCount = dataSet?.Tables?.Count;
                for (var i = 0; i < sheetsCount; i++)
                {
                    var table = dataSet?.Tables[i];
                    var name = table.TableName;
                    var isNewTable = tableDict.TryAdd(name, new ConcurrentDictionary<string, ExpandoObject>());
                    var rowsDict = tableDict[name];
                    Log.Information($"Processing '{name}' sheet (New: {isNewTable})");
                    var rows = table.GetRows<ExpandoObject>().ToList();
                    if (rows?.Count > 0)
                    {
                        foreach (var row in rows)
                        {
                            var id = row.Id();
                            object resultRow = null;
                            if (keepRight)
                            {
                                resultRow = rowsDict.AddOrUpdate(id, row, (key, existing) => row);
                            }
                            else if (keepLeft)
                            {
                                resultRow = rowsDict.AddOrUpdate(id, row, (key, existing) => existing);
                            }
                            else
                            {
                                var rowExists = rowsDict.TryGetValue(id, out var existingRow);
                                if (rowExists)
                                {
                                    var newProps = row.AllProps();
                                    var existingProps = rowsDict[row.Id()].AllProps();
                                    if (existingProps.Equals(newProps))
                                    {
                                        resultRow = row;
                                    }
                                    else
                                    {
                                        Log.Warning($"Keep row from (L)eft or (R)ight? (L / R)\nL: {existingProps}\nR: {newProps}");
                                        var choice = Console.ReadKey(true);
                                        Log.Information(choice.Key.ToString());
                                        if (choice.Key == ConsoleKey.R)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(id, row, (key, existing) => row);
                                        }
                                        else if (choice.Key == ConsoleKey.L)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(id, row, (key, existing) => existing);
                                        }
                                        else
                                        {
                                            Log.Warning("Invalid option!");
                                        }
                                    }
                                }
                                else
                                {
                                    if (rowsDict.TryAdd(id, row))
                                    {
                                        resultRow = row;
                                    }
                                }
                            }

                            Debug.Assert(resultRow != null, "Something went wrong during Merge operation!");
                        }
                    }
                }
            }

            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                var dataSet = new DataSet();
                foreach (var table in tableDict)
                {
                    var records = table.Value.Select(x => x.Value).ToList();
                    CheckDuplicates(records);
                    dataSet.Tables.Add(records.ToDataTable(table.Key));
                }

                workbook.Worksheets.Add(dataSet);
                workbook.SaveAs(output);
            }

            return 0;
        }

        public static int Diff(IEnumerable<string> inputs, string output)
        {
            var datasets = new Dictionary<string, Dictionary<string, IList<IDictionary<string, object>>>>();
            foreach (var input in inputs)
            {
                datasets.Add(input, input.GetData().ToExpandoProps());
            }

            var compareLogic = new CompareLogic
            {
                Config = new ComparisonConfig
                {
                    MaxDifferences = int.MaxValue,
                    IgnoreCollectionOrder = true,
                    TreatStringEmptyAndNullTheSame = true,
                    CaseSensitive = false
                }
            };

            // TODO: Handle more than 2 inputs?
            var comparison = compareLogic.Compare(datasets.FirstOrDefault().Value, datasets.LastOrDefault().Value);
            if (comparison.Differences.Count > 0)
            {
                Log.Warning(comparison.DifferencesString);
            }

            return 0;
        }

        private static void CheckDuplicates(IEnumerable<ExpandoObject> rows)
        {
            var dupRows = rows.GroupBy(x => x.AllProps(), StringComparer.OrdinalIgnoreCase)?.Where(g => g.Count() > 1).ToList();
            var dupKeys = rows.GroupBy(x => x.Id(), StringComparer.OrdinalIgnoreCase)?.Where(g => g.Count() > 1).ToList();
            DumpDuplicates(dupKeys, "Keys");
            DumpDuplicates(dupRows, "Rows");
        }

        private static void DumpDuplicates(List<IGrouping<string, ExpandoObject>> dups, string name)
        {
            if (dups?.Count > 0)
            {
                Log.Warning($"Duplicate {name} ({dups.Count}):");
                foreach (var row in dups)
                {
                    Log.Warning($"\t{row.Key} ({row.Count()})");
                }
            }
        }
    }
}
