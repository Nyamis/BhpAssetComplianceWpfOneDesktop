using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels
{
    public class MineComplianceBudgetMovementProduction
    {
        public DateTime Date { get; set; }
        public double LosColoradosOreGradeCutPercentage { get; set; }
        public double LosColoradosMillRecoveryPercentage { get; set; }
        public double LosColoradosMillFeedTonnes { get; set; }
        public double LosColoradosCuExMillTonnes { get; set; }
        public double LosColoradosRuntimePercentage { get; set; }
        public double LosColoradosHoursHours { get; set; }

        public double LagunaSecaOreGradeCutPercentage { get; set; }
        public double LagunaSecaMillRecoveryPercentage { get; set; }
        public double LagunaSecaMillFeedTonnes { get; set; }
        public double LagunaSecaCuExMillTonnes { get; set; }
        public double LagunaSecaRuntimePercentage { get; set; }
        public double LagunaSecaHoursHours { get; set; }

        public double LagunaSeca2OreGradeCutPercentage { get; set; }
        public double LagunaSeca2MillRecoveryPercentage { get; set; }
        public double LagunaSeca2MillFeedTonnes { get; set; }
        public double LagunaSeca2CuExMillTonnes { get; set; }
        public double LagunaSeca2RuntimePercentage { get; set; }
        public double LagunaSeca2HoursHours { get; set; }

        public double OxideOreToOlTonnes { get; set; }
        public double OxideCuCathodesTonnes { get; set; }

        public double SulphideLeachStackedMaterialFromMineTonnes { get; set; }
        public double SulphideLeachContractorsStackedMaterialFromStocksTonnes { get; set; }
        public double SulphideLeachMelStackedMaterialFromStocksTonnesTonnes { get; set; }
        public double SulphideLeachTotalStackedMaterialTonnes { get; set; }
        public double SulphideLeachCuCathodesTonnes { get; set; }
    }
}
