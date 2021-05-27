using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels
{
    public class MineComplianceBudgetPrincipal
    {
        public DateTime Date { get; set; }
        public double ExpitTonnes { get; set; }
        public double RehandlingTonnes { get; set; }
        public double MovementTonnes { get; set; }

        public double ShovelsUnits73Yd3 { get; set; }
        public double ShovelsAvailabilityPercentage { get; set; }
        public double ShovelsUtilizationPercentage { get; set; }
        public double ShovelsPerformanceTonnesPerHour { get; set; }
        public double ShovelsStandByHours { get; set; }
        public double ShovelsProductionTimeHours { get; set; }
        public double ShovelAvailableHoursHours { get; set; }
        public double ShovelHoursHours { get; set; }

        public double TrucksUnits { get; set; }
        public double TrucksAvailabilityPercentage { get; set; }
        public double TrucksUtilizationPercentage { get; set; }
        public double TrucksPerformanceTonnesPerDay { get; set; }
        public double TrucksStandByHours { get; set; }
        public double TrucksHoursHours { get; set; }
        public double TrucksProductionTimeHours { get; set; }
        public double TrucksAvailableHoursHours { get; set; }

        public double MillThroughputTonnes { get; set; }
        public double MillGradeCuPercentage { get; set; }
        public double MillRecoveryPercentage { get; set; }
        public double MillRehandlingPercentage { get; set; }

        public double OlThroughputTonnes { get; set; }
        public double OlGradeCuPercentage { get; set; }
        public double OlRecoveryPercentage { get; set; }
        public double OlCuSPercentage { get; set; }

        public double SlThroughputTonnes { get; set; }
        public double SlGradeCuPercentage { get; set; }
        public double SlRecoveryPercentage { get; set; }
        public double SlCuSPercentage { get; set; }

        public double MillProductionTonnes { get; set; }
        public double CathodesTonnes { get; set; }
        public double TotalProductionTonnes { get; set; }
    }
}
