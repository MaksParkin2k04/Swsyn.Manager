using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timesheets {

    /// <summary>
    /// Содержит информацию о задаче проекта
    /// </summary>
    public class ProjectTask {

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="project">Название проекта</param>
        /// <param name="typeService">Тип услуги</param>
        /// <param name="task">Название задачи</param>
        /// <param name="timeDays">Время потраченное на выполнение задачи</param>
        public ProjectTask(string project, string typeService, string task, IReadOnlyList<TimeDay> timeDays) {
            Project = project;
            TypeService = typeService;
            Task = task;
            TimeDays = timeDays;
        }

        /// <summary>
        /// Название проекта
        /// </summary>
        public string Project { get; private set; }

        /// <summary>
        /// Тип услуги
        /// </summary>
        public string TypeService { get; private set; }

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
        public double? GetSumBillable() {
            return TimeDays.Sum(t => t.Billable);
        }
    }
}
