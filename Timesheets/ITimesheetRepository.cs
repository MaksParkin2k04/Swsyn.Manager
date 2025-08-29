using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timesheets {
    public interface ITimesheetRepository {

        /// <summary>
        /// Возвращает коллекцию табелей учета рабочего времени для определенной недели
        /// </summary>
        /// <param name="mondayDate">Дата понедельника с которого начинается неделя</param>
        /// <returns>Коллекцию табелей учета рабочего времени</returns>
        Timesheet[] GetTimesheet(DateTime mondayDate);

        /// <summary>
        /// Сохраняет табель учета рабочего времени
        /// </summary>
        /// <param name="targetPath">Корневой директорий</param>
        /// <param name="projectName">Название проекта</param>
        /// <param name="timesheet">Табель учета рабочего времени</param>
        /// <param name="contractor">Исполнитель</param>
        /// <param name="customer">Заказчик</param>
        void WriteTimesheet(string targetPath, string projectName, Timesheet timesheet, string contractor, string customer);
    }
}
