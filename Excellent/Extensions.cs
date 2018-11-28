namespace Excellent
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using ExcelDataReader;
    using Serilog;

    public static class Extensions
    {
        public static DataSet GetSqlData(this string input, string connString)
        {
            try
            {
                var queryString = File.Exists(input.GetFullPath()) ? File.ReadAllText(input.GetFullPath()) : input;
                using (var adapter = new SqlDataAdapter(queryString, connString))
                {
                    var resultSet = new DataSet();
                    adapter.Fill(resultSet);
                    return resultSet;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, ex.Message);
            }

            return null;
        }

        public static DataSet GetExcelData(this string input)
        {
            try
            {
                using (var stream = File.Open(input, FileMode.Open, FileAccess.Read))
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
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
                        result.DataSetName = input;
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

        public static DataTable ToDataTable(this IList<ExpandoObject> items, string name)
        {
            if (items?.Count > 0)
            {
                var dt = new DataTable(name);
                foreach (var key in ((IDictionary<string, object>)items.FirstOrDefault()).Keys)
                {
                    dt.Columns.Add(key);
                }

                foreach (var item in items)
                {
                    dt.Rows.Add(((IDictionary<string, object>)item).Values.ToArray());
                }

                return dt;
            }

            return null;
        }

        public static T ToDictionary<T>(this DataRow row)
            where T : IDictionary<string, object>
        {
            var expandoDict = new ExpandoObject() as IDictionary<string, object>;
            foreach (DataColumn col in row.Table.Columns)
            {
                expandoDict.Add(col.ToString(), row[col.ColumnName].ToString());
            }

            return (T)expandoDict;
        }

        public static T ToDictionary<T>(this Item item)
            where T : IDictionary<string, object>
        {
            var expandoDict = new ExpandoObject() as IDictionary<string, object>;
            foreach (var prop in item.Props.Keys)
            {
                expandoDict.Add(prop, item.Props[prop]);
            }

            return (T)expandoDict;
        }

        public static bool ContainsIgnoreCase(this string item, string subString)
        {
            return item.IndexOf(subString, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string GetFullPath(this string file)
        {
            try
            {
                var value = Path.IsPathRooted(file) ? file : Path.Combine(Path.GetDirectoryName(new Uri(Assembly.GetCallingAssembly().CodeBase).LocalPath), file);
                return value;
            }
            catch
            {
                // Do nothing
            }

            return file;
        }
    }
}
