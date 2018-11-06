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
            var workbook = new Workbook(input.GetData());
            Log.Information($"Found {workbook.Sheets.Count} sheets\n");

            for (var i = 0; i < workbook.Sheets.Count; i++)
            {
                var sheet = workbook.Sheets[i];
                Log.Information($"Processing '{sheet.Name}' sheet");
                var items = sheet.Items;
                if (items?.Count > 0)
                {
                    Log.Information($"Row Count: {items.Count}");
                    CheckDuplicates(sheet);
                    var result = new StringBuilder();
                    result.AppendLine($"-- {sheet.Name}");
                    foreach (var row in items)
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
            var sheets = new ConcurrentDictionary<string, Worksheet>();
            foreach (var input in inputs)
            {
                var workbook = new Workbook(input.GetData());
                var sheetsCount = workbook.Sheets.Count;
                for (var i = 0; i < sheetsCount; i++)
                {
                    var workbookSheet = workbook.Sheets[i];
                    var isNew = sheets.TryAdd(workbookSheet.Name, workbookSheet);
                    var sheet = sheets[workbookSheet.Name];
                    Log.Information($"Processing '{workbookSheet.Name}' sheet (New: {isNew})");
                    var items = sheet.Items;
                    if (items?.Count > 0)
                    {
                        foreach (var item in items)
                        {
                            var id = item.Id;
                            object resultRow = null;
                            if (keepRight)
                            {
                                resultRow = sheet.AddOrUpdate(id, item, (key, existing) => item);
                            }
                            else if (keepLeft)
                            {
                                resultRow = sheet.AddOrUpdate(id, item, (key, existing) => existing);
                            }
                            else
                            {
                                var rowExists = sheet.ContainsItem(id);
                                if (rowExists)
                                {
                                    var newProps = item.FlattenValues();
                                    var existingProps = sheet.GetItem(item.Id).FlattenValues();
                                    if (existingProps.Equals(newProps, StringComparison.OrdinalIgnoreCase))
                                    {
                                        resultRow = item;
                                    }
                                    else
                                    {
                                        Log.Warning($"Keep row from (L)eft or (R)ight? (L / R)\nL: {existingProps}\nR: {newProps}");
                                        var choice = Console.ReadKey(true);
                                        Log.Information(choice.Key.ToString());
                                        if (choice.Key == ConsoleKey.R)
                                        {
                                            resultRow = sheet.AddOrUpdate(id, item, (key, existing) => item);
                                        }
                                        else if (choice.Key == ConsoleKey.L)
                                        {
                                            resultRow = sheet.AddOrUpdate(id, item, (key, existing) => existing);
                                        }
                                        else
                                        {
                                            Log.Warning("Invalid option!");
                                        }
                                    }
                                }
                                else
                                {
                                    resultRow = sheet.GetOrAdd(id, item);
                                }
                            }

                            Debug.Assert(resultRow != null, "Something went wrong during Merge operation!");
                        }
                    }
                }
            }

            Workbook.Save(output, sheets.Values.ToList());
            return 0;
        }

        public static int Diff(IEnumerable<string> inputs, string output)
        {
            //var datasets = new Dictionary<string, Dictionary<string, IList<IDictionary<string, object>>>>();
            //foreach (var input in inputs)
            //{
            //    datasets.Add(input, input.GetData().ToExpandoProps());
            //}

            //var compareLogic = new CompareLogic
            //{
            //    Config = new ComparisonConfig
            //    {
            //        MaxDifferences = int.MaxValue,
            //        IgnoreCollectionOrder = true,
            //        TreatStringEmptyAndNullTheSame = true,
            //        CaseSensitive = false
            //    }
            //};

            //// TODO: Handle more than 2 inputs?
            //var comparison = compareLogic.Compare(datasets.FirstOrDefault().Value, datasets.LastOrDefault().Value);
            //if (comparison.Differences.Count > 0)
            //{
            //    Log.Warning(comparison.DifferencesString);
            //}

            return 0;
        }

        private static void CheckDuplicates(Worksheet sheet)
        {
            var dupKeys = sheet.GetDuplicateItems(x => x.Id);
            DumpDuplicates(dupKeys, "Keys");

            var dupRows = sheet.GetDuplicateItems(x => x.FlattenValues());
            DumpDuplicates(dupRows, "Rows");
        }

        private static void DumpDuplicates(List<IGrouping<string, Item>> dups, string name)
        {
            if (dups?.Count > 0)
            {
                Log.Warning($"Duplicate {name} ({dups.Count}):");
                foreach (var dup in dups)
                {
                    Log.Warning($"\t{dup.Key} ({dup.Count()})");
                }
            }
        }
    }
}
