using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Swsyn.Manager.Models
{
    /// <summary>
    /// Содержит информацию о задачах выполненных за неделю
    /// </summary>
    public class TimesheetTable
    {

        /// <summary>
        /// Дата начала недели
        /// </summary>
        public DateTime StartDate { get; private set; }

        /// <summary>
        /// Выполненные задачи
        /// </summary>
        public IReadOnlyList<TaskRow> Tasks { get; private set; }

        /// <summary>
        /// Вычисляет оплачиваемое время потраченное на выполнение задач за определенный день
        /// </summary>
        /// <param name="date">Дата за которую производится вычисление</param>
        /// <returns>Время потраченное на выполнение задач</returns>
        public double? GetSumBillableFromDay(DateTime date)
        {
            double? result = null;
            foreach (var task in Tasks)
            {
                foreach (TimeDay timeDay in task.TimeDays)
                {
                    if (timeDay.Date == date)
                    {
                        if (timeDay.Billable != null)
                        {
                            result = (result == null) ? timeDay.Billable : result + timeDay.Billable;
                        }
                    }
                }
            }
            return result;
        }

        public double? GetSumBillableAll()
        {
            double? result = null;
            foreach (var task in Tasks)
            {
                foreach (TimeDay timeDay in task.TimeDays)
                {
                    if (timeDay.Billable != null)
                    {
                        result = (result == null) ? timeDay.Billable : result + timeDay.Billable;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="startDate">Дата начала недели</param>
        /// <param name="tasks">Выполненные задачи</param>
        public TimesheetTable(DateTime startDate, IReadOnlyList<TaskRow> tasks)
        {
            StartDate = startDate;
            Tasks = tasks;
        }
    }
}
