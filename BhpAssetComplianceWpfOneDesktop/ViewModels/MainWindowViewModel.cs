using System.Collections.ObjectModel;
using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Services;
using Prism.Commands;
using Prism.Mvvm;
using System.Windows.Input;
using BhpAssetComplianceWpfOneDesktop.ViewModels.DataTemplate;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        private string _myTitle = StringResources.ApplicationName;
        public string MyTitle
        {
            get { return _myTitle; }
            set { SetProperty(ref _myTitle, value); }
        }

        public ObservableCollection<RibbonButtonDataTemplateViewModel> OurRibbonItems { get; set; }
        public MainWindowViewModel(IAssetComplianceView assetComplianceView)
        {
            OurRibbonItems = new ObservableCollection<RibbonButtonDataTemplateViewModel>
            {
                new RibbonButtonDataTemplateViewModel(StringResources.MineSequence,assetComplianceView.MineSequenceView),
                new RibbonButtonDataTemplateViewModel(StringResources.MineCompliance,assetComplianceView.MineComplianceView),
                new RibbonButtonDataTemplateViewModel(StringResources.DepressurizationCompliance,assetComplianceView.DepressurizationComplianceView),
                new RibbonButtonDataTemplateViewModel(StringResources.Geotechnical,assetComplianceView.GeotechnicalView),
                new RibbonButtonDataTemplateViewModel(StringResources.QuartersReconciliationFactors,assetComplianceView.QuartersReconciliationFactorsView),
                new RibbonButtonDataTemplateViewModel(StringResources.ProcessCompliance,assetComplianceView.ProcessComplianceView),
                new RibbonButtonDataTemplateViewModel(StringResources.ConcentrateQuality,assetComplianceView.ConcentrateQualityView),
                new RibbonButtonDataTemplateViewModel(StringResources.HistoricalRecord,assetComplianceView.HistoricalRecordView),

            };
        }
    }

}
