using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.ProcessComplianceModels
{
    public class ProcessComplianceRecovery
    {
        public double RecGlobalBudget { get; set; }
        public double RecGlobalActual { get; set; }
        public double RecGlobalMD { get; set; }
        public string Phase { get; set; }
        public double RecoveryBudget { get; set; }
        public double RecoveryActual { get; set; }
        public double FeedCuBudget { get; set; }
        public double FeedCuActual { get; set; }
    }
}
