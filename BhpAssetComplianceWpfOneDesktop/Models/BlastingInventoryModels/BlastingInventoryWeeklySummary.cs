using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.BlastingInventoryModels
{
    public class BlastingInventoryWeeklySummary
    {
        public string Week { get; set; }
        public double? AvgWeeklyEscondidaTonnes { get; set; }
        public double? SumWeeklyEscondidaEvents { get; set; }
        public double? TargetWeeklyEscondidaEvents { get; set; }
        public double? AvgWeeklyEscondidaNorteTonnes { get; set; }
        public double? SumWeeklyEscondidaNorteEvents { get; set; }
        public double? TargetWeeklyEscondidaNorteEvents { get; set; }
    }
}

