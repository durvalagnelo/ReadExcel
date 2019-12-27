using ReadExcel.Interfaces;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    public class SpreadSheet : ISpreadSheet
    {
        public SLDocument SLDocument { get; private set; }

        public SpreadSheet(string filePath, string sheetName)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                SLDocument = new SLDocument(fileStream, sheetName);
            }
        }

        public int CountRows()
        {
            return SLDocument.GetWorksheetStatistics().EndRowIndex;
        }

        public int MaxColumns()
        {
            return SLDocument.GetWorksheetStatistics().EndColumnIndex;
        }

        public SpreadSheetDataFrame ReadAllLines()
        {
            List<ReadOnlyCollection<string>> __return = new List<ReadOnlyCollection<string>>();

            SLWorksheetStatistics sLWorksheetStatistics = SLDocument.GetWorksheetStatistics();
            for (int row = 1; row < sLWorksheetStatistics.EndRowIndex; row++)
            {
                ReadOnlyCollection<string> columns = ReadLine(row);
                __return.Add(columns);
            }

            return new SpreadSheetDataFrame(__return);
        }

        public ReadOnlyCollection<string> ReadLine(int row)
        {
            List<string> __return = new List<string>();

            SLWorksheetStatistics sLWorksheetStatistics = SLDocument.GetWorksheetStatistics();

            for (int column = 0; column <= sLWorksheetStatistics.EndColumnIndex; column++)
            {
                string value = SLDocument.GetCellValueAsString(row, column);
                __return.Add(value);
            }

            return __return.AsReadOnly();

        }
    }
}
