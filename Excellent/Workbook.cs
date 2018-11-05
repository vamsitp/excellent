namespace Excellent
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Dynamic;
    using System.Linq;

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
    }

    public class Worksheet
    {
        public Worksheet(DataTable dataTable)
        {
            this.Name = dataTable.TableName;
            this.Items = dataTable.AsEnumerable().Select(r => new Item(r)).ToList();
        }

        public string Name { get; set; }

        public IList<Item> Items { get; set; }

        public List<IGrouping<string, Item>> GetDuplicates(Func<Item, string> groupSelector)
        {
            var dups = this.Items.GroupBy(groupSelector, StringComparer.OrdinalIgnoreCase)?.Where(g => g.Count() > 1).ToList();
            return dups;
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
    }
}
