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

    using SmartFormat;

    public static class Extensions
    {
        public static string Id(this object obj)
        {
            var id = Smart.Format(ConfigurationManager.AppSettings["PrimaryKey"], obj);
            return id;
        }

        public static string AllProps(this ExpandoObject obj, string delimiter = " | ", bool namesOnly = false)
        {
            var format = string.Join(delimiter, obj.AllPropsList(namesOnly));
            var result = format.Trim(delimiter.Trim().ToCharArray());
            return result;
        }

        public static List<object> AllPropsList(this ExpandoObject obj, bool namesOnly = false)
        {
            var props = new List<object>();
            var dict = obj as IDictionary<string, object>;
            if (dict == null)
            {
                dict = obj.GetType().GetProperties().ToDictionary(x => x.Name, x => x.GetValue(obj));
            }

            foreach (var prop in dict)
            {
                props.Add(namesOnly ? prop.Key : prop.Value);
            }

            return props;
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

        public static Dictionary<string, IList<IDictionary<string, object>>> ToExpandoProps(this DataSet ds)
        {
            var expandoDict = new Dictionary<string, IList<IDictionary<string, object>>>();
            foreach (DataTable dt in ds.Tables)
            {
                var expandoList = dt.GetList<IDictionary<string, object>>();
                expandoDict.Add(dt.TableName, expandoList);
            }

            return expandoDict;
        }

        public static Dictionary<string, List<ExpandoObject>> ToExpandoDict(this DataSet ds)
        {
            var expandoDict = new Dictionary<string, List<ExpandoObject>>();
            foreach (DataTable dt in ds.Tables)
            {
                var expandoList = dt.ToExpandoList();
                expandoDict.Add(dt.TableName, expandoList);
            }

            return expandoDict;
        }

        public static List<ExpandoObject> ToExpandoList(this DataTable dt)
        {
            var expandoList = new List<ExpandoObject>();
            foreach (DataRow row in dt.Rows)
            {
                var expandoDict = ToExpandoObject(dt, row);
                expandoList.Add(expandoDict);
            }

            return expandoList;
        }

        public static ExpandoObject ToExpandoObject(DataTable dt, DataRow row)
        {
            var expandoDict = new ExpandoObject() as IDictionary<string, object>;
            foreach (DataColumn col in dt.Columns)
            {
                expandoDict.Add(col.ToString(), row[col.ColumnName].ToString());
            }

            return (ExpandoObject)expandoDict;
        }

        public static bool TryAdd(this DataTableCollection dtc, DataTable dt)
        {
            if (dtc.Contains(dt.TableName))
            {
                return false;
            }

            dtc.Add(dt);
            return true;
        }

        public static bool AddOrUpdate(this DataRowCollection drc, object key, DataRow addValue, Func<string, DataRow, DataRow> updateValueFactory)
        {
            var dr = drc.Find(key);
            if (dr != null)
            {
                dr.Delete();
                drc.Add(updateValueFactory);
                return false;
            }

            drc.Add(addValue);
            return true;
        }
    }
}
