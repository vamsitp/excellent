namespace Excellent
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Text;

    using ClosedXML.Excel;

    using CommandLine;

    using CsvHelper;
    using CsvHelper.Excel;

    using Serilog;

    using SmartFormat;
    using SmartFormat.Core.Settings;

    internal class Program
    {
        private readonly static string OutputFormat = ConfigurationManager.AppSettings["TransformFormat"];

        private static void Main(string[] args)
        {
            SetLogger();
            SetSmartFormatting();
            var result = Parser.Default.ParseArguments<TransformOptions, MergeOptions, DiffOptions>(args).MapResult((TransformOptions opts) => Transform(opts), (MergeOptions opts) => Merge(opts), (DiffOptions opts) => Diff(opts), errs => HandleParseErrors(errs?.ToList()));
            if (result == 0)
            {
                Log.Information("Done!");
            }

            Log.CloseAndFlush();
            if (Debugger.IsAttached)
            {
                Console.ReadLine();
            }
        }

        private static int Transform(TransformOptions opts)
        {
            var input = opts.Input;
            Log.Information($"Processing '{input}'");

            var dataSet = input.GetData();
            var sheetsCount = dataSet?.Tables?.Count;
            Log.Information($"Found {sheetsCount} sheets\n");

            var outputFile = !string.IsNullOrWhiteSpace(opts.Output) ? opts.Output : (Path.GetFileNameWithoutExtension(input) + ".txt");
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
                        var val = Smart.Format(OutputFormat, row);
                        result.AppendLine(val);
                    }

                    result.AppendLine();
                    Log.Information($"Writing output to '{outputFile}'\n");
                    if (i == 0)
                    {
                        File.WriteAllText(outputFile, result.ToString());
                    }
                    else
                    {
                        File.AppendAllText(outputFile, result.ToString());
                    }
                }
            }

            return 0;
        }

        private static int Merge(MergeOptions opts)
        {
            var inputs = opts.Inputs;
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
                    Log.Information($"Processing '{name}' sheet");
                    var rows = table.GetRows<ExpandoObject>().ToList();
                    if (rows?.Count > 0)
                    {
                        foreach (var row in rows)
                        {
                            var id = row.Id();
                            object resultRow = null;
                            if (opts.KeepRight)
                            {
                                resultRow = rowsDict.AddOrUpdate(id, row, (key, existingVal) => row);
                            }
                            else if (opts.KeepLeft)
                            {
                                resultRow = rowsDict.AddOrUpdate(id, row, (key, existingRow) => existingRow);
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
                                        if (choice.Key == ConsoleKey.R)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(id, row, (key, existingVal) => row);
                                        }
                                        else if (choice.Key == ConsoleKey.L)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(id, row, (key, existingVal) => row);
                                        }
                                        else
                                        {
                                            Console.WriteLine("Invalid option!");
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

            var outputFile = !string.IsNullOrWhiteSpace(opts.Output) ? opts.Output : (string.Join("_", inputs.Select(Path.GetFileNameWithoutExtension)) + "_Merged.xlsx");
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                foreach (var table in tableDict)
                {
                    var worksheet = workbook.AddWorksheet(table.Key);
                    using (var writer = new CsvWriter(new ExcelSerializer(worksheet)))
                    {
                        var records = table.Value.Select(x =>
                        {
                            var val = x.Value;
                            return val;
                        }).ToList();

                        writer.WriteRecords(records);
                    }
                }

                workbook.SaveAs(outputFile);
            }

            return 0;
        }

        private static int Diff(DiffOptions opts)
        {
            return 0;
        }

        private static int HandleParseErrors(List<Error> errs)
        {
            if (errs.Count > 1)
            {
                errs.ToList().ForEach(e => Log.Error(e.ToString()));
            }

            return errs.Count;
        }

        private static void SetSmartFormatting()
        {
            Smart.Default.Settings.ConvertCharacterStringLiterals = false;
            Smart.Default.Settings.CaseSensitivity = CaseSensitivityType.CaseInsensitive;
            Smart.Default.Settings.FormatErrorAction = ErrorAction.Ignore;
        }

        private static void CheckDuplicates(IEnumerable<ExpandoObject> rows)
        {
            var dupRows = rows.GroupBy(x => x.AllProps())?.Count(g => g.Count() > 1);
            var dupKeys = rows.GroupBy(x => x.Id())?.Count(g => g.Count() > 1);
            Log.Warning($"Duplicates: Key = {dupKeys ?? 0} | Values = {dupRows ?? 0}");
        }

        private static void SetLogger()
        {
            Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.Console(outputTemplate: "[{Level:u3}] {Message}{NewLine}").WriteTo.RollingFile("Excellent_{Date}.log", outputTemplate: "{Timestamp:dd-MMM-yyyy HH:mm:ss} | [{Level}] {Message}{NewLine}{Exception}").Enrich.FromLogContext().CreateLogger();
        }
    }
}
