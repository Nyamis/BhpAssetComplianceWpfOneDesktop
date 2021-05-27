using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels
{
    public class MineComplianceRealLoadingFc
    {
        public string Name { get; set; }
        public double Units { get; set; }
        public double AvailabilityPercentage { get; set; }
        public double UtilizationPercentage { get; set; }
        public double TotalHoursHours { get; set; }
        public double AvailableHoursHours { get; set; }
        public double EquipmentScheduledDowntimeHours { get; set; }
        public double EquipmentNonScheduledDowntimeHours { get; set; }
        public double ProcessScheduledDowntimeHours { get; set; }
        public double ProcessNonScheduledDowntimeHours { get; set; }
        public double StandByHours { get; set; }
        public double HangTimeHours { get; set; }
        public double ProductionTimeHours { get; set; }
        public double PerformanceTonnesPerHour { get; set; }
        public double TotalTonnesTonnes { get; set; }
    }
}
