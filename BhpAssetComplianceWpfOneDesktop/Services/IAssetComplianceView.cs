using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Services
{
    public interface IAssetComplianceView
    {

        void MineComplianceView();
        void GeotechnicalView();
        void MineSequenceView();
        void DepressurizationComplianceView();
        void ProcessComplianceView();
        void ConcentrateQualityView();
        void QuartersReconciliationFactorsView();
        void HistoricalRecordView();
    }
}
