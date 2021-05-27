using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.ProcessComplianceModels
{
    public class ProcessComplianceOreToMill
    {
        public double SpiGlobalBudget { get; set; }
        public double SpiGlobalActual { get; set; }
        public string Phase { get; set; }
        public double OretoMillBudget { get; set; }
        public double OretoMillActual { get; set; }
        public double HardnessBudget { get; set; }
        public double HardnessActual { get; set; }
    }
}
