using System.Collections.Generic;
using BhpAssetComplianceWpfOneDesktop.Models.QuartersReconciliationFactorsModels;

namespace BhpAssetComplianceWpfOneDesktop.Engines
{
    public struct QuarterReconciliationFactors
    {
        public List<QuartersReconciliationFactorsF0> F0 { get; set; }
        public List<QuartersReconciliationFactorsF1> F1 { get; set; }
        public List<QuartersReconciliationFactorsF2> F2 { get; set; }
        public List<QuartersReconciliationFactorsF3> F3 { get; set; }
    }
}