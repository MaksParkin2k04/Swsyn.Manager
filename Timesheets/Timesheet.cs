using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timesheets {

    /// <summary>
    /// Содержит информацию о задачах выполненных за неделю
    /// </summary>
    public class Timesheet {

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="startDate">Дата начала недели</param>
        /// <param name="tasks">Выполненные задачи</param>
        public Timesheet(DateTime startDate, IReadOnlyList<ProjectTask> tasks) {
            StartDate = startDate;
            Tasks = tasks;
        }

        /// <summary>
        /// Дата начала недели
        /// </summary>
        public DateTime StartDate { get; private set; }

        /// <summary>
        /// Выполненные задачи
        /// </summary>
        public IReadOnlyList<ProjectTask> Tasks { get; private set; }

        /// <summary>
        /// Вычисляет оплачиваемое время потраченное на выполнение задач за определенный день
        /// </summary>
        /// <param name="date">Дата за которую производится вычисление</param>
        /// <returns>Время потраченное на выполнение задач</returns>
        public double? GetSumBillableFromDay(DateTime date) {
            double? result = null;
            foreach (ProjectTask task in Tasks) {
                foreach (TimeDay timeDay in task.TimeDays) {
                    if (timeDay.Date == date) {
                        if (timeDay.Billable != null) {
                            result = result == null ? timeDay.Billable : result + timeDay.Billable;
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Вычисляет общее оплачиваемое время потраченное на выполнение задач за неделю
        /// </summary>
        /// <returns>Время потраченное на выполнение задач</returns>
        public double? GetSumBillableAll() {
            double? result = null;
            foreach (ProjectTask task in Tasks) {
                foreach (TimeDay timeDay in task.TimeDays) {
                    if (timeDay.Billable != null) {
                        result = result == null ? timeDay.Billable : result + timeDay.Billable;
                    }
                }
            }
            return result;
        }
    }
}
