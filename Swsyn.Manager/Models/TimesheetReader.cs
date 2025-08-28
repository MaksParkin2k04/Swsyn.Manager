//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace Swsyn.Manager.Models
//{
//    public class TimesheetReader : IDisposable
//    {

//        /// <summary>
//        /// Массив названий ячеек содержащих информацию о времени выполнения задачи
//        /// </summary>
//        private static string[] TaskTimeDayCellNames = new string[] { "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R" };

//        private SpreadsheetDocument document;
//        private WorkbookPart workbookPart;
//        private WorksheetPart worksheetPart;
//        private SheetData sheetData;

//        /// <summary>
//        /// Конструктор
//        /// </summary>
//        /// <param name="filePath">Путь к файлу</param>
//        public TimesheetReader(string filePath)
//        {
//            document = SpreadsheetDocument.Open(filePath, false);
//            workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
//            worksheetPart = workbookPart.WorksheetParts.First();
//            sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
//        }

//        /// <summary>
//        /// Считывает последнюю таблицу в документе. 
//        /// При считывании учитываются как видимые так и скрытые строки задач.
//        /// </summary>
//        /// <returns>Данные из последней таблицы</returns>
//        public TimesheetTable? ReadLastTable()
//        {
//            TimesheetTable? table = null;

//            // Перебираем все строки начиная с конца
//            foreach (Row row in sheetData.Elements<Row>().Reverse())
//            {
//                // Ищем начало таблицы
//                if (IsStartTable(row) == true)
//                {
//                    table = ReadTable(row);
//                    break;
//                }
//            }

//            return table;
//        }

//        /// <summary>
//        /// Считывает таблицу начинающуюся со строки rowIndex.
//        /// </summary>
//        /// <param name="rowIndex">Строка начала таблицы</param>
//        /// <returns>Таблица</returns>
//        private TimesheetTable ReadTable(Row firstRow)
//        {
//            // Получаем дату начала таблицы
//            DateTime? startDate = GetTableStartDate(firstRow);

//            // Вычисляем сколько строк нужно пропустить 
//            int skipRow = (int)firstRow.RowIndex!.Value + 3;

//            List<TaskRow> taskRows = new List<TaskRow>();

//            // Перебираем строки содержащие информацию о задачах
//            foreach (Row row in sheetData.Elements<Row>().Skip(skipRow))
//            {
//                // Если найден конец таблицы выходим из цикла
//                if (IsEndTable(row) == true) { break; }

//                // Получаем задачу из строки
//                TaskRow taskRow = ReadTaskRow(row, startDate!.Value);
//                taskRows.Add(taskRow);
//            }


//            return new TimesheetTable(startDate!.Value, taskRows);
//        }

//        /// <summary>
//        /// Считывает информацию о задаче
//        /// </summary>
//        /// <param name="row">Строка содержащая информацию о задаче</param>
//        /// <param name="startDate">Дата первого дня недели</param>
//        /// <returns>Задача</returns>
//        private TaskRow ReadTaskRow(Row row, DateTime startDate)
//        {
//            string projectName = GetProjectName(row) ?? string.Empty;
//            string typeOfServise = Get
//            string taskName = GetTaskName(row) ?? string.Empty;
//            List<TimeDay> timeDays = new List<TimeDay>();

//            int indexDay = 0;
//            for (int i = 0; i < TaskTimeDayCellNames.Length; i = i + 2)
//            {
//                DateTime taskDate = startDate.AddDays(indexDay++);

//                Cell? reportedCell = GetCell(row, TaskTimeDayCellNames[i]);
//                Cell? billableCell = GetCell(row, TaskTimeDayCellNames[i + 1]);
//                double? reported = ReadDouble(reportedCell);
//                double? billable = ReadDouble(billableCell);

//                TimeDay timeDay = new TimeDay(taskDate, reported, billable);
//                timeDays.Add(timeDay);
//            }

//            return new TaskRow(projectName, typeOfServise, taskName, timeDays);
//        }

//        /// <summary>
//        /// Считывает имя проекта
//        /// </summary>
//        /// <param name="row">Строка содержащая задачу</param>
//        /// <returns>Имя проекта</returns>
//        private string? GetProjectName(Row row)
//        {
//            Cell? cell = GetCell(row, "A");
//            if (cell == null) { return null; }
//            return ReadSharedString(cell);
//        }



//        /// <summary>
//        /// Считывает имя задачи
//        /// </summary>
//        /// <param name="row">Строка содержащая задачу</param>
//        /// <returns>Имя задачи</returns>
//        private string? GetTaskName(Row row)
//        {
//            Cell? cell = GetCell(row, "B");
//            if (cell == null) { return null; }
//            return ReadSharedString(cell);
//        }

//        /// <summary>
//        /// Проверяет является ли строка началом таблицы
//        /// </summary>
//        /// <param name="row">Проверяемая строка</param>
//        /// <returns>True, если строка является началом таблицы, иначе false</returns>
//        private bool IsStartTable(Row row)
//        {
//            Cell? cell = GetCell(row, "A");
//            if (cell == null) { return false; }
//            string? text = ReadSharedString(cell);
//            return text == "ДАТА НАЧАЛА НЕДЕЛИ";
//        }

//        /// <summary>
//        /// Проверяет является ли строка последней строкой таблицы
//        /// </summary>
//        /// <param name="row">Проверяемая строка</param>
//        /// <returns>True, если строка является последней строкой таблицы, иначе false</returns>
//        private bool IsEndTable(Row row)
//        {
//            Cell? cell = GetCell(row, "B");
//            if (cell == null) { return false; }
//            string? text = ReadSharedString(cell);
//            return text == "ИТОГО ЧАСОВ";
//        }

//        /// <summary>
//        /// Возвращает дату начала таблицы
//        /// </summary>
//        /// <param name="row">Первая строка таблицы</param>
//        /// <returns></returns>
//        private DateTime? GetTableStartDate(Row row)
//        {
//            Cell? cell = GetCell(row, "S");
//            return ReadDate(cell);
//        }

//        /// <summary>
//        /// Возвращает ячейку строки по названию колонки
//        /// </summary>
//        /// <param name="row">Строка содержащая ячейку</param>
//        /// <param name="columnName">Название колонки</param>
//        /// <returns>Найденная ячейка или null</returns>
//        private Cell? GetCell(Row row, string columnName)
//        {
//            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0);
//        }

//        /// <summary>
//        /// Считывает содержимое ячейки как число
//        /// </summary>
//        /// <param name="cell">Ячейка</param>
//        /// <returns>Число или null</returns>
//        private double? ReadDouble(Cell? cell)
//        {
//            double? result = null;

//            if (cell == null || cell.InnerText == null) { return null; }

//            if (cell.DataType == null || cell.DataType == CellValues.Number)
//            {
//                if (double.TryParse(cell.InnerText, CultureInfo.InvariantCulture, out double value))
//                {
//                    result = value;
//                }
//            }
//            return result;
//        }

//        /// <summary>
//        /// Считывает содержимое ячейки как дату
//        /// </summary>
//        /// <param name="cell">Ячейка</param>
//        /// <returns>Дата или null</returns>
//        private DateTime? ReadDate(Cell? cell)
//        {
//            if (cell != null && cell.DataType == null && int.TryParse(cell.InnerText, out int countDay))
//            {
//                return DateTime.FromOADate(countDay);
//            }
//            return null;
//        }

//        /// <summary>
//        /// Считывает содержимое ячейки как строку
//        /// </summary>
//        /// <param name="cell">Ячейка</param>
//        /// <returns>Строка или null</returns>
//        private string? ReadSharedString(Cell? cell)
//        {
//            string? result = null;

//            if (cell == null || cell.InnerText == null) { return null; }

//            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
//            {
//                if (Int32.TryParse(cell.InnerText, out int id))
//                {
//                    SharedStringItem? item = GetSharedStringItemById(id);
//                    result = item?.Text?.Text;
//                }
//            }
//            return result;
//        }

//        /// <summary>
//        /// Возвращает строку из общих ресурсов по ее идентификатору
//        /// </summary>
//        /// <param name="id">Идентификатор строки</param>
//        /// <returns></returns>
//        public SharedStringItem? GetSharedStringItemById(int id)
//        {
//            return workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAt(id);
//        }

//        /// <summary>
//        /// Освобождает ресурсы
//        /// </summary>
//        public void Dispose()
//        {
//            if (document != null)
//            {
//                document.Dispose();
//            }
//        }
//    }
//}
