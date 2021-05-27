using System.Collections.ObjectModel;
using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Services;
using Prism.Commands;
using Prism.Mvvm;
using System.Windows.Input;
using BhpAssetComplianceWpfOneDesktop.Constants;
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
                new RibbonButtonDataTemplateViewModel(StringResources.MineSequence,assetComplianceView.MineSequenceView,IconKeys.MineSequence),
                new RibbonButtonDataTemplateViewModel(StringResources.MineCompliance,assetComplianceView.MineComplianceView,IconKeys.MineMovement),
                new RibbonButtonDataTemplateViewModel(StringResources.DepressurizationCompliance,assetComplianceView.DepressurizationComplianceView,IconKeys.Depressurization),
                new RibbonButtonDataTemplateViewModel(StringResources.Geotechnical,assetComplianceView.GeotechnicalView,IconKeys.Geotechnics),
                new RibbonButtonDataTemplateViewModel(StringResources.QuartersReconciliationFactors,assetComplianceView.QuartersReconciliationFactorsView,IconKeys.ReconciliationFactors),
                new RibbonButtonDataTemplateViewModel(StringResources.ProcessCompliance,assetComplianceView.ProcessComplianceView,IconKeys.ProcessCompliance),
                new RibbonButtonDataTemplateViewModel(StringResources.ConcentrateQuality,assetComplianceView.ConcentrateQualityView,IconKeys.ConcentrateQuality),
                new RibbonButtonDataTemplateViewModel(StringResources.HistoricalRecord,assetComplianceView.HistoricalRecordView,IconKeys.KvdSummary),
                new RibbonButtonDataTemplateViewModel(StringResources.Repository,assetComplianceView.RepositoryView,IconKeys.Repository),
            };
        }
    }

}
