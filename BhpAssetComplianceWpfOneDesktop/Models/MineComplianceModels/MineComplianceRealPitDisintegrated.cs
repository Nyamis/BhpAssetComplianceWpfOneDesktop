using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels
{
    public class MineComplianceRealPitDisintegrated
    {
        public double ExpitEsTonnes { get; set; }
        public double ExpitEnTonnes { get; set; }
        public double TotalExpitTonnes { get; set; }

        public double MillRehandlingTonnes { get; set; }
        public double OlRehandlingTonnes { get; set; }
        public double SlRehandlingTonnes { get; set; }
        public double OtherRehandlingTonnes { get; set; }
        public double TotalRehandlingTonnes { get; set; }

        public double TotalMovementTonnes { get; set; }

        public double RehandlingTotalTonnes { get; set; }
        public double MovementTotalTonnes { get; set; }

        public double TotalTonnes { get; set; }
    }
}
