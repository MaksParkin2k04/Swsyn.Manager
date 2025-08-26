using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Swsyn.Manager.Models
{
    public class TaskRow
    {
        /// <summary>
        /// Название проекта
        /// </summary>
        public string Project { get; private set; }

        /// <summary>
        /// Вид услуги
        /// </summary>
        public string TypeOfService { get; private set; }

        /// <summary>
        /// Название задачи
        /// </summary>
        public string Task { get; private set; }

        /// <summary>
        /// Время потраченное на выполнение задачи
        /// </summary>
        public IReadOnlyList<TimeDay> TimeDays { get; private set; }

        /// <summary>
        /// Вычисляет оплачиваемое время потраченное на выполнение задачи за все дни
        /// </summary>
        /// <returns>Время потраченное на выполнение задачи</returns>
        public double? GetSumBillable()
        {
            return TimeDays.Sum(t => t.Billable);
        }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="project">Название проекта</param>
        /// <param name="typeOfServise">Время потраченное на выполнение задачи</param>
        /// <param name="task">Название задачи</param>
        /// <param name="timeDays">Время потраченное на выполнение задачи</param>
        public TaskRow(string project, string typeOfServise, string task, IReadOnlyList<TimeDay> timeDays)
        {
            Project = project;
            TypeOfService = typeOfServise;
            Task = task;
            TimeDays = timeDays;
        }
    }
}
