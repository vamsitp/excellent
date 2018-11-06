namespace Excellent
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Dynamic;
    using System.Linq;

    using ClosedXML.Excel;

    using SmartFormat;

    public class Workbook
    {
        public Workbook(DataSet dataSet)
        {
            this.Name = dataSet.DataSetName;
            this.Sheets = dataSet.Tables.Cast<DataTable>().Select(t => new Worksheet(t)).ToList();
        }

        public string Name { get; set; }

        public IList<Worksheet> Sheets { get; set; }

        public bool ContainsSheet(string name, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            return this.ContainsSheet(x => x.Name.Equals(name, comparison));
        }

        public bool ContainsSheet(Func<Worksheet, bool> condition)
        {
            return this.Sheets.Any(condition);
        }

        public Worksheet GetSheet(string name, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            return this.Sheets.SingleOrDefault(x => x.Name.Equals(name, comparison));
        }

        public IEnumerable<Worksheet> GetSheets(Func<Worksheet, bool> condition)
        {
            return this.Sheets.Where(condition);
        }

        public bool Save(string output)
        {
            return Save(output, this);
        }

        public static bool Save(string output, Workbook workbook)
        {
            return Save(output, workbook.Sheets);
        }

        public static bool Save(string output, IList<Worksheet> sheets)
        {
            using (var wb = new XLWorkbook(XLEventTracking.Disabled))
            {
                foreach (var sheet in sheets)
                {
                    sheet.AddToWorkbook(wb);
                }

                wb.SaveAs(output);
            }

            return true;
        }
    }

    public class Worksheet
    {
        private const string FontFamily = "Segoe UI";

        public Worksheet(DataTable dataTable)
        {
            this.Name = dataTable.TableName;
            this.Items = dataTable.AsEnumerable().Select(r => new Item(r)).ToList();
        }

        public string Name { get; set; }

        public IList<Item> Items { get; set; }

        public bool ContainsItem(string id, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            return this.ContainsItem(x => x.Id.Equals(id, comparison));
        }

        public bool ContainsItem(Func<Item, bool> condition)
        {
            return this.Items.Any(condition);
        }

        public Item GetItem(string id, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            return this.GetItem(x => x.Id.Equals(id, comparison));
        }

        public Item GetItem(Func<Item, bool> condition)
        {
            return this.Items.FirstOrDefault(condition); // TOEDO: SingleOrDefault
        }

        public IEnumerable<Item> GetItems(Func<Item, bool> condition)
        {
            return this.Items.Where(condition);
        }

        public List<IGrouping<string, Item>> GetDuplicateItems(Func<Item, string> groupSelector)
        {
            var dups = this.Items.GroupBy(groupSelector, StringComparer.OrdinalIgnoreCase)?.Where(g => g.Count() > 1).ToList();
            return dups;
        }

        public bool TryAdd(Item item)
        {
            if (this.ContainsItem(item.Id))
            {
                return false;
            }

            this.Items.Add(item);
            return true;
        }

        public Item GetOrAdd(string id, Item addValue)
        {
            var items = this.GetItems(x => x.Id.Equals(id, StringComparison.Ordinal)).ToList();
            if (items?.Count > 0)
            {
                return items.SingleOrDefault();
            }

            this.Items.Add(addValue);
            return addValue;
        }

        public bool AddOrUpdate(string id, Item addValue, Func<string, Item, Item> updateValueFactory)
        {
            // TODO: GetItem
            var items = this.GetItems(x => x.Id.Equals(id, StringComparison.Ordinal)).ToList();
            if (items?.Count > 0)
            {
                items.ForEach(x => x.Props = updateValueFactory(x.Id, x).Props);
                return false;
            }

            this.Items.Add(addValue);
            return true;
        }

        public DataTable ToDataTable()
        {
            if (this.Items?.Count > 0)
            {
                var dt = new DataTable(this.Name);
                foreach (var key in this.Items.FirstOrDefault().Props.Keys)
                {
                    dt.Columns.Add(key);
                }

                foreach (var item in this.Items)
                {
                    dt.Rows.Add(item.Props.Values.ToArray());
                }

                return dt;
            }

            return null;
        }

        public IXLWorksheet AddToWorkbook(IXLWorkbook workbook)
        {
            var ws = workbook.AddWorksheet(this.Name);
            var cols = this.Items.FirstOrDefault().Props.Keys;
            var header = ws.Cell(1, 1).InsertData(cols, true);
            ws.Cell(2, 1).InsertData(this.ToDataTable());
            ws.RangeUsed().SetAutoFilter();
            ws.Style.Font.SetFontName(FontFamily);
            ws.Style.Font.SetFontSize(10);
            header.Style.Font.Bold = true;
            ws.Column(1).AdjustToContents();
            ws.Column(1).AddConditionalFormat().WhenIsDuplicate().Font.SetFontColor(XLColor.Red);
            ws.Column(2).AddConditionalFormat().WhenIsDuplicate().Font.SetFontColor(XLColor.BrickRed);
            ws.Cell(1, 2).SetActive();
            ws.SheetView.Freeze(1, 2);
            return ws;
        }
    }

    public class Item : DynamicObject
    {
        public string Id { get; set; }

        public IDictionary<string, object> Props { get; set; }

        public Item(DataRow dataRow)
        {
            this.Props = dataRow.ToDictionary<IDictionary<string, object>>();
            this.Id = Smart.Format(ConfigurationManager.AppSettings["PrimaryKey"], this.Props);
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (this.Props.ContainsKey(binder.Name))
            {
                result = this.Props[binder.Name];
                return true;
            }
            else
            {
                result = "Invalid Property!";
                return false;
            }
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            this.Props[binder.Name] = value;
            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            dynamic method = this.Props[binder.Name];
            result = method(args[0].ToString(), args[1].ToString());
            return true;
        }

        public string FlattenValues(string delimiter = " | ")
        {
            var result = string.Join(delimiter, this.Props.Select(p => this.Props[p.Key]));
            return result;
        }

        public string FlattenNames(string delimiter = " | ")
        {
            var result = string.Join(delimiter, this.Props.Select(p => p.Key));
            return result;
        }

        public override bool Equals(object obj)
        {
            var item = obj as Item;
            return item != null &&
                   this.FlattenNames().Equals(item.FlattenNames(), StringComparison.OrdinalIgnoreCase) &&
                   this.FlattenValues().Equals(item.FlattenValues(), StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            var hashCode = -681290639;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(this.FlattenNames());
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(this.FlattenValues());
            return hashCode;
        }
    }
}
