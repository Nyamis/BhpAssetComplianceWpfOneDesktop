using System;
using System.Collections.Generic;
using BhpAssetComplianceWpfOneDesktop.Resources;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using BhpAssetComplianceWpfOneDesktop.Constants;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class HistoricalRecordViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.HistoricalRecord;
        protected override string MyPosterIcon { get; set; } = IconKeys.KvdSummary;

        private string _myLastRefreshValues;
        public string MyLastRefreshValues
        {
            get { return _myLastRefreshValues; }
            set { SetProperty(ref _myLastRefreshValues, value); }
        }

        private DateTime _myDateActual;
        public DateTime MyDateActual
        {
            get { return _myDateActual; }
            set { SetProperty(ref _myDateActual, value); }
        }

        private string _myComment;
        public string MyComment
        {
            get { return _myComment; }
            set { SetProperty(ref _myComment, value); }
        }

        private List<String> _Areas;
        public List<String> Areas
        {
            get
            {
                return new List<string>() {
                    StringResources.All,
                    StringResources.MineSequence,
                    StringResources.MineCompliance,
                    StringResources.DepressurizationCompliance,
                    StringResources.Geotechnical,
                    StringResources.QuartersReconciliationFactors,
                    StringResources.ProcessCompliance,
                    StringResources.ConcentrateQuality,
                    StringResources.BlastingInventory,
                    StringResources.Other };
            }
            set
            {
                _Areas = value;
            }
        }

        private string _myArea;
        public string MyArea
        {
            get { return _myArea; }
            set { SetProperty(ref _myArea, value); }
        }

        public DelegateCommand LoadRecordCommand { get; private set; }

        public HistoricalRecordViewModel()
        {
            MyDateActual = DateTime.Now;
            LoadRecordCommand = new DelegateCommand(LoadNewRecord);
        }

        private void LoadNewRecord()
        {
            if (MyComment != null & MyArea != null)
            {
                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.HistoricalRecordExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);

                if (loadFileInfo.Exists)
                {
                    var package = new ExcelPackage(loadFileInfo);
                    var worksheet = package.Workbook.Worksheets[HistoricalRecordConstants.HistoricalRecordWorksheet];

                    if (worksheet != null)
                    {
                        try
                        {
                            var openWriteCheck = File.OpenWrite(loadFilePath);
                            openWriteCheck.Close();

                            var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, MyDateActual.Day, 00, 00, 00);
                            var lastRow = worksheet.Dimension.End.Row + 1;

                            worksheet.Cells[lastRow, 1].Value = newDate;
                            worksheet.Cells[lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            worksheet.Cells[lastRow, 2].Value = MyArea;
                            worksheet.Cells[lastRow, 3].Value = MyComment;

                            byte[] fileText = package.GetAsByteArray();
                            File.WriteAllBytes(loadFilePath, fileText);
                            MyLastRefreshValues = $"{StringResources.Updated}: {DateTime.Now}";

                            MyComment = null;
                            MyArea = null;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, StringResources.UploadError);
                        }
                    }
                    else
                    {
                        var wrongFileMessage = $"{StringResources.WorksheetNotExist} {loadFilePath} {StringResources.IsTheRightOne}";
                        MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                    }                   
                }
                else
                {
                    var wrongFileMessage = $"{StringResources.WorksheetNotExist} {loadFilePath} {StringResources.ExistsOrNotSelect}";
                    MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                }
            }
            else
            {
                MessageBox.Show(StringResources.SelectAreaInputComment, StringResources.UploadError);
            }

        }
    }
}
