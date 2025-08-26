using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Swsyn.Manager.Models
{
    public class Helper
    {
        public static Row? GetRow(SheetData sheetData, uint rowIndex)
        {
            return sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex == rowIndex);
        }

        public static Cell? GetCell(SheetData sheetData, string columnName, uint rowIndex)
        {
            Row? row = GetRow(sheetData, rowIndex);
            if (row == null) { return null; }
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        }

        public static Cell? GetCell(Row? row, string columnName)
        {
            if (row == null) { return null; }
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0);
        }
    }
}
