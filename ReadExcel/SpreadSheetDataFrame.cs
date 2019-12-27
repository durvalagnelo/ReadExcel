using ReadExcel.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    public class SpreadSheetDataFrame : ReadOnlyCollection<ReadOnlyCollection<string>>, ISpreadSheetDataFrame
    {
        public SpreadSheetDataFrame(IList<ReadOnlyCollection<string>> list) : base(list)
        {
        }

        public string GetString(int row, int column)
        {
            return this[row][column];
        }
        public DateTime GetDateTime(int row, int column)
        {
            return Convert.ToDateTime(this[row][column]);
        }
        public int GetInt(int row, int column)
        {
            return Convert.ToInt32(this[row][column]);
        }
        public decimal? GetDecimalNullable(int row, int column)
        {
            string value = this[row][column];

            if (String.IsNullOrEmpty(value.Trim()))
            {
                return null;
            }

            decimal dValue = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            return dValue;
        }
        public decimal GetDecimal(int row, int column)
        {
            decimal? valueNullable = GetDecimalNullable(row, column);
            return valueNullable.HasValue ? valueNullable.Value : 0m;
        }
    }
}
