namespace Excellent
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Configuration;
    using System.Data;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;

    using ExcelDataReader;
    using Newtonsoft.Json;
    using SmartFormat;

    public static class Extensions
    {
        public static string Id(this object obj)
        {
            var id = Smart.Format(ConfigurationManager.AppSettings["PrimaryKey"], obj);
            return id;
        }

        public static string AllProps(this ExpandoObject obj)
        {
            var format = new StringBuilder();
            var dict = obj as IDictionary<string, object>;
            if (dict == null)
            {
                dict = obj.GetType().GetProperties().ToDictionary(x => x.Name, x => x.GetValue(obj));
            }

            foreach (var prop in dict)
            {
                format.Append("{" + prop.Key + "} | ");
            }

            var result = Smart.Format(format.ToString(), obj).TrimEnd(' ', '|');
            return result;
        }

        public static IEnumerable<T> GetRows<T>(this string input, int sheetIndex = 0)
        {
            var dataSet = GetData(input);
            return GetRows<T>(dataSet, sheetIndex);
        }

        public static IEnumerable<T> GetRows<T>(this DataSet dataSet, int sheetIndex)
        {
            var dataTable = dataSet?.Tables[sheetIndex];
            return GetRows<T>(dataTable);
        }

        public static IEnumerable<T> GetRows<T>(this DataTable dataTable)
        {
            return GetList<T>(dataTable);
        }

        public static DataSet GetData(this string input)
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

        public static List<T> GetList<T>(this DataTable dt)
        {
            var json = JsonConvert.SerializeObject(dt);
            var escs = (ConfigurationManager.GetSection("escapeChars") as NameValueCollection ?? throw new ConfigurationErrorsException("escapeChars"));
            foreach (var esc in escs.AllKeys)
            {
                json = json.Replace(esc, escs[esc]);
            }

            var result = JsonConvert.DeserializeObject<List<T>>(json);
            return result;
        }
    }
}
