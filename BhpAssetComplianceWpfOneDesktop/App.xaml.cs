﻿using System;
using BhpAssetComplianceWpfOneDesktop.Services;
using BhpAssetComplianceWpfOneDesktop.Utility;
using BhpAssetComplianceWpfOneDesktop.Views;
using Prism.Ioc;
using System.Windows;
using System.Windows.Media;

namespace BhpAssetComplianceWpfOneDesktop
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App
    {
        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<IAssetComplianceView, AssetComplianceView>();
            containerRegistry.RegisterForNavigation<GeotechnicalView>(ViewNames.GeotechnicalView);
            containerRegistry.RegisterForNavigation<MineComplianceView>(ViewNames.MineComplianceView);
            containerRegistry.RegisterForNavigation<MineSequenceView>(ViewNames.MineSequenceView);
            containerRegistry.RegisterForNavigation<DepressurizationComplianceView>(ViewNames.DepressurizationComplianceView);
            containerRegistry.RegisterForNavigation<ProcessComplianceView>(ViewNames.ProcessComplianceView);
            containerRegistry.RegisterForNavigation<ConcentrateQualityView>(ViewNames.ConcentrateQualityView);
            containerRegistry.RegisterForNavigation<QuartersReconciliationFactorsView>(ViewNames.QuartersReconciliationFactorsView);           
            containerRegistry.RegisterForNavigation<BlastingInventoryView>(ViewNames.BlastingInventoryView);
            containerRegistry.RegisterForNavigation<HistoricalRecordView>(ViewNames.HistoricalRecordView);
            containerRegistry.RegisterForNavigation<RepositoryView>(ViewNames.RepositoryView);
        }

        protected override void OnInitialized()
        {
            base.OnInitialized();
            Container.Resolve<IAssetComplianceView>().MineSequenceView();
        }
    }
}
