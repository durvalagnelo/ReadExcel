using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace ReadExcel.Interfaces
{
    public interface ISpreadSheet
    {
        int CountRows();

        int MaxColumns();

        SpreadSheetDataFrame ReadAllLines();

        ReadOnlyCollection<string> ReadLine(int row);
    }
}
