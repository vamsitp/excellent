namespace Excellent
{
    using System.Collections.Generic;

    class ExcelRow
    {
        public string Id => $"{this.ResourceId}_{this.ResourceSet}";

        public string ResourceId { get; set; }

        public string English { get; set; }

        public string French { get; set; }

        public string Spanish { get; set; }

        public string ResourceSet { get; set; }

        public override bool Equals(object obj)
        {
            var row = obj as ExcelRow;
            return row != null &&
                   this.ResourceId == row.ResourceId &&
                   this.English == row.English &&
                   this.ResourceSet == row.ResourceSet;
        }

        public override int GetHashCode()
        {
            var hashCode = 272569366;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(this.ResourceId);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(this.English);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(this.ResourceSet);
            return hashCode;
        }

        public override string ToString()
        {
            return $"{this.ResourceId} | {this.English} | {this.ResourceSet}";
        }
    }
}
