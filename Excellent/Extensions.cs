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

    using ExcelDataReader;

    using Newtonsoft.Json;

    public static class Extensions
    {
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
    }
}
