namespace Excellent
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;

    using ClosedXML.Excel;

    using CommandLine;

    using CsvHelper;
    using CsvHelper.Excel;

    using ExcelDataReader;

    using Serilog;

    internal class Program
    {
        private readonly static string OutputFormat = ConfigurationManager.AppSettings["TransformFormat"];

        private static void Main(string[] args)
        {
            SetLogger();
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

        private static int HandleParseErrors(List<Error> errs)
        {
            if (errs.Count > 1)
            {
                errs.ToList().ForEach(e => Log.Error(e.ToString()));
            }

            return errs.Count;
        }

        private static int Diff(DiffOptions opts)
        {
            return 0;
        }

        private static int Merge(MergeOptions opts)
        {
            var inputs = opts.Inputs;
            Log.Information($"Processing '{string.Join(", ", inputs)}'");
            var tableDict = new ConcurrentDictionary<string, ConcurrentDictionary<string, ExcelRow>>();
            foreach (var input in inputs)
            {
                var dataSet = GetData(input);
                var sheetsCount = dataSet?.Tables?.Count;
                for (var i = 0; i < sheetsCount; i++)
                {
                    var table = dataSet?.Tables[i];
                    var name = table.TableName;
                    var isNewTable = tableDict.TryAdd(name, new ConcurrentDictionary<string, ExcelRow>());
                    var rowsDict = tableDict[name];
                    Log.Information($"Processing '{name}' sheet");
                    var rows = GetRows<ExcelRow>(table).ToList();
                    if (rows?.Count > 0)
                    {
                        foreach (var row in rows)
                        {
                            ExcelRow resultRow = null;
                            if (opts.KeepRight)
                            {
                                resultRow = rowsDict.AddOrUpdate(row.Id, row, (key, existingVal) => row);
                            }
                            else if (opts.KeepLeft)
                            {
                                resultRow = rowsDict.AddOrUpdate(row.Id, row, (key, existingRow) => existingRow);
                            }
                            else
                            {
                                var rowExists = rowsDict.TryGetValue(row.Id, out var existingRow);
                                if (rowExists)
                                {
                                    if (existingRow.Equals(row))
                                    {
                                        resultRow = row;
                                    }
                                    else
                                    {
                                        Log.Warning($"Keep row from (L)eft or (R)ight? (L / R)\n'L: {rowsDict[row.Id]}'\n'R: {row}'");
                                        var choice = Console.ReadKey(true);
                                        if (choice.Key == ConsoleKey.R)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(row.Id, row, (key, existingVal) => row);
                                        }
                                        else if (choice.Key == ConsoleKey.L)
                                        {
                                            resultRow = rowsDict.AddOrUpdate(row.Id, row, (key, existingVal) => row);
                                        }
                                        else
                                        {
                                            Console.WriteLine("Invalid option!");
                                        }
                                    }
                                }
                                else
                                {
                                    if (rowsDict.TryAdd(row.Id, row))
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
                        writer.WriteRecords(table.Value.Select(x => x.Value));
                    }
                }

                workbook.SaveAs(outputFile);
            }

            return 0;
        }

        private static int Transform(TransformOptions opts)
        {
            var input = opts.Input;
            Log.Information($"Processing '{input}'");

            var dataSet = GetData(input);
            var sheetsCount = dataSet?.Tables?.Count;
            Log.Information($"Found {sheetsCount} sheets\n");

            var outputFile = !string.IsNullOrWhiteSpace(opts.Output) ? opts.Output : (Path.GetFileNameWithoutExtension(input) + ".sql");
            for (var i = 0; i < sheetsCount; i++)
            {
                var table = dataSet?.Tables[i];
                Log.Information($"Processing '{table.TableName}' sheet");
                var rows = GetRows<ExcelRow>(table).ToList();
                if (rows?.Count > 0)
                {
                    Log.Information($"Row Count: {rows.Count}");
                    CheckDuplicates(rows);
                    var sql = new StringBuilder();
                    sql.AppendLine($"-- {table.TableName}");
                    foreach (var row in rows)
                    {
                        sql.AppendLine(string.Format(OutputFormat, row.ResourceId, row.English, row.French, row.Spanish, row.ResourceSet));
                    }

                    sql.AppendLine();
                    Log.Information($"Writing output to '{outputFile}'\n");
                    var result = sql.ToString().Replace("'", "''");
                    if (i == 0)
                    {
                        File.WriteAllText(outputFile, result);
                    }
                    else
                    {
                        File.AppendAllText(outputFile, result);
                    }
                }
            }

            return 0;
        }

        private static void CheckDuplicates(IEnumerable<ExcelRow> rows)
        {
            var dupRows = rows.GroupBy(x => x.ResourceId + x.English + x.ResourceSet)?.Count(g => g.Count() > 1);
            var dupKeys = rows.GroupBy(x => x.ResourceId)?.Count(g => g.Count() > 1);
            var dupValues = rows.GroupBy(x => x.English)?.Count(g => g.Count() > 1);
            Log.Warning($"Duplicates: Entire-row = {dupRows ?? 0} | ResourceId = {dupKeys ?? 0} | English = {dupValues ?? 0}");

            var frenchDefaults = rows.Count(x => x.French.StartsWith("fr-CA"));
            var spanishDefaults = rows.Count(x => x.Spanish.StartsWith("es-MX"));
            Log.Warning($"Dev Defaults: French = {frenchDefaults} | Spanish = {spanishDefaults}");
        }

        private static void SetLogger()
        {
            Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.Console(outputTemplate: "[{Level:u3}] {Message}{NewLine}").WriteTo.RollingFile("Excellent_{Date}.log", outputTemplate: "{Timestamp:dd-MMM-yyyy HH:mm:ss} | [{Level}] {Message}{NewLine}{Exception}").Enrich.FromLogContext().CreateLogger();
        }

        public static IEnumerable<T> GetRows<T>(string input, int sheetIndex = 0)
        {
            var dataSet = GetData(input);
            return GetRows<T>(dataSet, sheetIndex);
        }

        private static IEnumerable<T> GetRows<T>(DataSet dataSet, int sheetIndex)
        {
            var dataTable = dataSet?.Tables[sheetIndex];
            return GetRows<T>(dataTable);
        }

        private static IEnumerable<T> GetRows<T>(DataTable dataTable)
        {
            return GetList<T>(dataTable);
        }

        public static DataSet GetData(string input)
        {
            try
            {
                using (var stream = File.Open(input, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                // reader.GetDouble(0);
                            }
                        }
                        while (reader.NextResult());

                        var options = new ExcelDataSetConfiguration
                        {
                            UseColumnDataType = true,
                            ConfigureDataTable = tableReader => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true,
                                FilterRow = rowReader => true
                            }
                        };

                        var result = reader.AsDataSet(options);
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return null;
        }

        private static List<T> GetList<T>(DataTable dt)
        {
            var data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                var item = GetItem<T>(row);
                data.Add(item);
            }

            return data;
        }

        private static T GetItem<T>(DataRow dr)
        {
            var temp = typeof(T);
            var obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (var pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                    {
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    }
                }
            }

            return obj;
        }
    }
}
