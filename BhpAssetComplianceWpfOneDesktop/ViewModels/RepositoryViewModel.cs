using BhpAssetComplianceWpfOneDesktop.Constants;
using BhpAssetComplianceWpfOneDesktop.Resources;
using Microsoft.Win32;
using Prism.Commands;
using System.IO;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class RepositoryViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.Repository;
        protected override string MyPosterIcon { get; set; } = IconKeys.Poster;

        private string _myMineSequenceExcelFilePath;
        public string MyMineSequenceExcelFilePath
        {
            get { return _myMineSequenceExcelFilePath; }
            set { SetProperty(ref _myMineSequenceExcelFilePath, value); }
        }

        private string _myMineSequenceCSVFilePath;
        public string MyMineSequenceCSVFilePath
        {
            get { return _myMineSequenceCSVFilePath; }
            set { SetProperty(ref _myMineSequenceCSVFilePath, value); }
        }

        private string _myMineComplianceExcelFilePath;
        public string MyMineComplianceExcelFilePath
        {
            get { return _myMineComplianceExcelFilePath; }
            set { SetProperty(ref _myMineComplianceExcelFilePath, value); }
        }

        private string _myDepressurizationComplianceExcelFilePath;
        public string MyDepressurizationComplianceExcelFilePath
        {
            get { return _myDepressurizationComplianceExcelFilePath; }
            set { SetProperty(ref _myDepressurizationComplianceExcelFilePath, value); }
        }

        private string _myDepressurizationComplianceCSVFilePath;
        public string MyDepressurizationComplianceCSVFilePath
        {
            get { return _myDepressurizationComplianceCSVFilePath; }
            set { SetProperty(ref _myDepressurizationComplianceCSVFilePath, value); }
        }

        private string _myGeotechnicalNotesExcelFilePath;
        public string MyGeotechnicalNotesExcelFilePath
        {
            get { return _myGeotechnicalNotesExcelFilePath; }
            set { SetProperty(ref _myGeotechnicalNotesExcelFilePath, value); }
        }

        private string _myGeotechnicalNotesCSVFilePath;
        public string MyGeotechnicalNotesCSVFilePath
        {
            get { return _myGeotechnicalNotesCSVFilePath; }
            set { SetProperty(ref _myGeotechnicalNotesCSVFilePath, value); }
        }

        private string _myQuartersReconciliationFactorsExcelFilePath;
        public string MyQuartersReconciliationFactorsExcelFilePath
        {
            get { return _myQuartersReconciliationFactorsExcelFilePath; }
            set { SetProperty(ref _myQuartersReconciliationFactorsExcelFilePath, value); }
        }

        private string _myProcessComplianceExcelFilePath;
        public string MyProcessComplianceExcelFilePath
        {
            get { return _myProcessComplianceExcelFilePath; }
            set { SetProperty(ref _myProcessComplianceExcelFilePath, value); }
        }

        private string _myConcentrateQualityExcelFilePath;
        public string MyConcentrateQualityExcelFilePath
        {
            get { return _myConcentrateQualityExcelFilePath; }
            set { SetProperty(ref _myConcentrateQualityExcelFilePath, value); }
        }

        private string _myBlastingInventoryExcelFilePath;
        public string MyBlastingInventoryExcelFilePath
        {
            get { return _myBlastingInventoryExcelFilePath; }
            set { SetProperty(ref _myBlastingInventoryExcelFilePath, value); }
        }

        private string _myBlastingInventoryCSVFilePath;
        public string MyBlastingInventoryCSVFilePath
        {
            get { return _myBlastingInventoryCSVFilePath; }
            set { SetProperty(ref _myBlastingInventoryCSVFilePath, value); }
        }


        private string _myHistoricalRecordExcelFilePath;
        public string MyHistoricalRecordExcelFilePath
        {
            get { return _myHistoricalRecordExcelFilePath; }
            set { SetProperty(ref _myHistoricalRecordExcelFilePath, value); }
        }

        public DelegateCommand SelectMineSequenceDataExcelFileCommand { get; set; }
        public DelegateCommand SelectMineSequenceDataCSVFileCommand { get; set; }
        public DelegateCommand SelectMineComplianceDataExcelFileCommand { get; set; }
        public DelegateCommand SelectDepressurizationComplianceDataExcelFileCommand { get; set; }
        public DelegateCommand SelectDepressurizationComplianceDataCSVFileCommand { get; set; }
        public DelegateCommand SelectGeotechnicalNotesDataExcelFileCommand { get; set; }
        public DelegateCommand SelectGeotechnicalNotesDataCSVFileCommand { get; set; }
        public DelegateCommand SelectQuartersReconciliationFactorsDataExcelFileCommand { get; set; }
        public DelegateCommand SelectProcessComplianceDataExcelFileCommand { get; set; }
        public DelegateCommand SelectConcentrateQualityDataExcelFileCommand { get; set; }
        public DelegateCommand SelectBlastingInventoryDataExcelFileCommand { get; set; }
        public DelegateCommand SelectBlastingInventoryDataCSVFileCommand { get; set; }
        public DelegateCommand SelectHistoricalRecordDataExcelFileCommand { get; set; }

        public RepositoryViewModel()
        {
            MyMineSequenceExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceExcelFilePath;
            MyMineSequenceCSVFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceCSVFilePath;
            MyMineComplianceExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineComplianceExcelFilePath;
            MyDepressurizationComplianceExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceExcelFilePath;
            MyDepressurizationComplianceCSVFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceCSVFilePath;
            MyGeotechnicalNotesExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesExcelFilePath;
            MyGeotechnicalNotesCSVFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesCSVFilePath;
            MyQuartersReconciliationFactorsExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.QuartersReconciliationFactorsExcelFilePath;
            MyProcessComplianceExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ProcessComplianceExcelFilePath;
            MyConcentrateQualityExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ConcentrateQualityExcelFilePath;
            MyBlastingInventoryExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.BlastingInventoryExcelFilePath;
            MyBlastingInventoryCSVFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.BlastingInventoryCSVFilePath;
            MyHistoricalRecordExcelFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.HistoricalRecordExcelFilePath;
            SelectMineSequenceDataExcelFileCommand = new DelegateCommand(SelectMineSequenceDataExcelFile);
            SelectMineSequenceDataCSVFileCommand = new DelegateCommand(SelectMineSequenceDataCSVFile);
            SelectMineComplianceDataExcelFileCommand = new DelegateCommand(SelectMineComplianceDataExcelFile);
            SelectDepressurizationComplianceDataExcelFileCommand = new DelegateCommand(SelectDepressurizationComplianceDataExcelFile);
            SelectDepressurizationComplianceDataCSVFileCommand = new DelegateCommand(SelectDepressurizationComplianceDataCSVFile);
            SelectGeotechnicalNotesDataExcelFileCommand = new DelegateCommand(SelectGeotechnicalNotesDataExcelFile);
            SelectGeotechnicalNotesDataCSVFileCommand = new DelegateCommand(SelectGeotechnicalNotesDataCSVFile);
            SelectQuartersReconciliationFactorsDataExcelFileCommand = new DelegateCommand(SelectQuartersReconciliationFactorsDataExcelFile);
            SelectProcessComplianceDataExcelFileCommand = new DelegateCommand(SelectProcessComplianceDataExcelFile);
            SelectConcentrateQualityDataExcelFileCommand = new DelegateCommand(SelectConcentrateQualityDataExcelFile);
            SelectBlastingInventoryDataExcelFileCommand = new DelegateCommand(SelectBlastingInventoryDataExcelFile);
            SelectBlastingInventoryDataCSVFileCommand = new DelegateCommand(SelectBlastingInventoryDataCSVFile);
            SelectHistoricalRecordDataExcelFileCommand = new DelegateCommand(SelectHistoricalRecordDataExcelFile);
        }       

        private void SelectMineSequenceDataExcelFile()
        {          
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyMineSequenceExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceExcelFilePath = MyMineSequenceExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectMineSequenceDataCSVFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "CSV file (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyMineSequenceCSVFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceCSVFilePath = MyMineSequenceCSVFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectMineComplianceDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyMineComplianceExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineComplianceExcelFilePath = MyMineComplianceExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectDepressurizationComplianceDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyDepressurizationComplianceExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceExcelFilePath = MyDepressurizationComplianceExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectDepressurizationComplianceDataCSVFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "CSV file (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyDepressurizationComplianceCSVFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceCSVFilePath = MyDepressurizationComplianceCSVFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectGeotechnicalNotesDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyGeotechnicalNotesExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesExcelFilePath = MyGeotechnicalNotesExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectGeotechnicalNotesDataCSVFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "CSV file (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyGeotechnicalNotesCSVFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesCSVFilePath = MyGeotechnicalNotesCSVFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectQuartersReconciliationFactorsDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyQuartersReconciliationFactorsExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.QuartersReconciliationFactorsExcelFilePath = MyQuartersReconciliationFactorsExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectProcessComplianceDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyProcessComplianceExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ProcessComplianceExcelFilePath = MyProcessComplianceExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectConcentrateQualityDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyConcentrateQualityExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ConcentrateQualityExcelFilePath = MyConcentrateQualityExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectBlastingInventoryDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyBlastingInventoryExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.BlastingInventoryExcelFilePath = MyBlastingInventoryExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectBlastingInventoryDataCSVFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "CSV file (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyBlastingInventoryCSVFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.BlastingInventoryCSVFilePath = MyBlastingInventoryCSVFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }

        private void SelectHistoricalRecordDataExcelFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyHistoricalRecordExcelFilePath = new FileInfo(openFileDialog.FileName).FullName;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.HistoricalRecordExcelFilePath = MyHistoricalRecordExcelFilePath;
                BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.Save();
            }
        }
    }
}
