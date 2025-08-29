
namespace Timesheets.OpenXml {
    public class TimesheetRepository : ITimesheetRepository {

        public TimesheetRepository(string dataDirectory) {
            this.dataDirectory = dataDirectory;
        }

        private readonly string dataDirectory;

        public Timesheet[] GetTimesheet(DateTime startDate) {
            List<Timesheet> tables = new List<Timesheet>();

            string[] files = GetDataFiles(startDate);
            foreach (string file in files) {
                using (TimesheetReader reader = new TimesheetReader(file)) {
                    Timesheet? table = reader.ReadTable(startDate);

                    // Если отчет о указанной неделе отсутствует
                    if (table == null) {
                        string currentTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                        string fileName = Path.GetFileName(file);
                        throw new Exception($"[{currentTime}]  Week report {startDate.ToString("dd.MM.yyyy")} was not found for {fileName}.");
                    }

                    tables.Add(table);
                }
            }

            return tables.ToArray();
        }

        public void WriteTimesheet(string targetPath, string projectName, Timesheet timesheet, string contractor, string customer) {
            string fileName = $"Timesheet_{projectName}_{timesheet.StartDate.ToString("ddMMyyyy")}.xlsx";
            string writePath = Path.Combine(targetPath, fileName);

            TimesheetWriter.Write(writePath, timesheet, contractor, customer);
        }

        private string[] GetDataFiles(DateTime specifyPeriod) {
            string year = specifyPeriod.Year.ToString();
            string searchPath = Path.Combine(dataDirectory, year);
            string searchPattern = $"*_{year}.xlsx";
            return Directory.GetFiles(searchPath, searchPattern);
        }


    }
}
