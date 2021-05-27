using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.ProcessComplianceModels
{
    public class ProcessComplianceOLAP
    {
        public string FeedGrade { get; set; }
        public double Budget { get; set; }
        public double Actual { get; set; }
        public double Compliance { get; set; }
        public string Distribution { get; set; }
        public double DistributionBudget { get; set; }
        public double DistributionActual { get; set; }
    }
}
