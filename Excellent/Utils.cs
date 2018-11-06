namespace Excellent
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;

    using Serilog;

    using SmartFormat;

    public static class Utils
    {
        public readonly static StringComparison IgnoreCase = bool.TryParse(ConfigurationManager.AppSettings[nameof(IgnoreCase)], out var ignoreCase) && ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

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
                        var val = Smart.Format(outputFormat, row.Props);
                        Debug.Assert(!val.Equals(outputFormat), $"'{Path.GetFileNameWithoutExtension(workbook.Name)} - {sheet.Name}''s '{row.Id}' row could not be transformed!");
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
                    Log.Information($"Processing '{Path.GetFileNameWithoutExtension(input)}' - '{workbookSheet.Name}' sheet (New: {isNew})");
                    var items = sheet.Items.Union(workbookSheet.Items).ToList();
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
                                    if (existingProps.Equals(newProps, Utils.IgnoreCase))
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
            var workbooks = new List<Workbook>();
            foreach (var input in inputs)
            {
                workbooks.Add(new Workbook(input.GetData()));
            }

            //// TODO: Handle more than 2 inputs?
            var first = workbooks.FirstOrDefault();
            var last = workbooks.LastOrDefault();
            var firstSheets = first.Sheets;
            var lastSheets = last.Sheets;
            for (var i = 0; i < firstSheets.Count; i++)
            {
                var firstSheet = firstSheets[i];
                var lastSheet = lastSheets[i];

                var firstSheetItems = firstSheets[i].Items;
                var lastSheetItems = lastSheets[i].Items;

                var firstOnlyItems = firstSheetItems.Except(lastSheetItems).Select(x => x.FlattenValues()).ToList();
                var lastOnlyItems = lastSheetItems.Except(firstSheetItems).Select(x => x.FlattenValues()).ToList();
                var matches = lastSheetItems.Intersect(firstSheetItems).ToList();
                Log.Warning($"'{Path.GetFileNameWithoutExtension(first.Name)}' - '{firstSheet.Name}': {firstSheetItems.Count}");
                Log.Warning($"'{Path.GetFileNameWithoutExtension(last.Name)}' - '{lastSheet.Name}': {lastSheetItems.Count}");
                Log.Warning($"Matches = {matches.Count}");

                if (firstOnlyItems?.Count > 0)
                {
                    Log.Warning($"{firstOnlyItems?.Count} items in '{Path.GetFileNameWithoutExtension(first.Name)}' not in '{Path.GetFileNameWithoutExtension(last.Name)}': \n\t{string.Join(Environment.NewLine + "\t", firstOnlyItems)}\n");
                }

                if (lastOnlyItems?.Count > 0)
                {
                    Log.Warning($"{lastOnlyItems?.Count} items in '{Path.GetFileNameWithoutExtension(last.Name)}' not in '{Path.GetFileNameWithoutExtension(first.Name)}': \n\t{string.Join(Environment.NewLine + "\t", lastOnlyItems)}\n");
                }
            }

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
