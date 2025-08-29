using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timesheets {

    /// <summary>
    /// Содержит данные о времени потраченном на выполнении задачи в определенную дату.
    /// </summary>
    public class TimeDay {

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="date">Дата выполнения задачи</param>
        /// <param name="reported">Время потраченное на выполнение задачи</param>
        /// <param name="billable">Оплачиваемое время потраченное на выполнение задачи</param>
        public TimeDay(DateTime date, double? reported, double? billable) {
            Date = date;
            Reported = reported;
            Billable = billable;
        }

        /// <summary>
        /// Дата выполнения задачи
        /// </summary>
        public DateTime Date { get; private set; }

        /// <summary>
        /// Время потраченное на выполнение задачи
        /// </summary>
        public double? Reported { get; private set; }

        /// <summary>
        /// Оплачиваемое время потраченное на выполнение задачи
        /// </summary>
        public double? Billable { get; private set; }
    }
}
