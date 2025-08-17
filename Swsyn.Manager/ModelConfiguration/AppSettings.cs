using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Swsyn.Manager.ModelConfiguration
{
    public class AppSettings
    {
        public string SourcePath { get; set; } = string.Empty;
        public Dictionary<string, ProjectOption> Projects { get; set; }
        public string[]? Include { get; set; }
    }
}
