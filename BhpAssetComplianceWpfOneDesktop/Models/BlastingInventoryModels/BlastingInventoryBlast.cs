using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.BlastingInventoryModels
{
    public class BlastingInventoryBlast
    {
        public DateTime Date { get; set; }
        public string Place { get; set; }
        public double? SulphideTotalTonnes { get; set; }
        public double? OthersTotalTonnes { get; set; }
        public double? DayBlastTonnes { get; set; }
        public double? Events { get; set; }
        public double? WeeklyEvents { get; set; }
        public double? TargetEvents { get; set; }
    }
}
