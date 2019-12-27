using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcel.Interfaces
{
    public interface ISpreadSheetDataFrame
    {
        string GetString(int row, int column);
        DateTime GetDateTime(int row, int column);
        int GetInt(int row, int column);
        decimal? GetDecimalNullable(int row, int column);
        decimal GetDecimal(int row, int column);
    }
}
