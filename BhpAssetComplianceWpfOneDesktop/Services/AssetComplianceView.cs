using BhpAssetComplianceWpfOneDesktop.Utility;
using Prism.Regions;

namespace BhpAssetComplianceWpfOneDesktop.Services
{
    public class AssetComplianceView : IAssetComplianceView
    {
        private readonly IRegionManager _regionManager;

        public AssetComplianceView(IRegionManager regionManager)
        {
            _regionManager = regionManager;
        }

        public void MineComplianceView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.MineComplianceView);
        }

        public void GeotechnicalView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.GeotechnicalView);
        }

        public void MineSequenceView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.MineSequenceView);
        }

        public void DepressurizationComplianceView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.DepressurizationComplianceView);
        }

        public void ProcessComplianceView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.ProcessComplianceView);
        }

        public void ConcentrateQualityView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.ConcentrateQualityView);
        }

        public void QuartersReconciliationFactorsView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.QuartersReconciliationFactorsView);
        }

        public void HistoricalRecordView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.HistoricalRecordView);
        }

        public void RepositoryView()
        {
            _regionManager.RequestNavigate(RegionNames.MainRegion, ViewNames.RepositoryView);
        }


    }
}