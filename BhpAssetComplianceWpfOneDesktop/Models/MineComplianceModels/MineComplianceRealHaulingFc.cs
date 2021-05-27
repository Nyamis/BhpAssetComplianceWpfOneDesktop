using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels
{
    public class MineComplianceRealHaulingFc
    {
        public string Name { get; set; }
        public double Units { get; set; }
        public double MechanicalAvailabilityPercentage { get; set; }
        public double PhysicalAvailabilityPercentage { get; set; }
        public double UtilizationPercentage { get; set; }
        public double TotalHoursHours { get; set; }
        public double AvailableHoursHours { get; set; }
        public double EquipmentScheduledDowntimeHours { get; set; }
        public double EquipmentNonScheduledDowntimeHours { get; set; }
        public double ProcessScheduledDowntimeHours { get; set; }
        public double ProcessNonScheduledDowntimeHours { get; set; }
        public double StandByHours { get; set; }
        public double QueueTimeHours { get; set; }
        public double ProductionTimeHoursHours { get; set; }
        public double PerformanceTonnesPerHour { get; set; }
        public double CycleTimeHours { get; set; }
        public double TotalTonnesTonnes { get; set; }
    }
}
