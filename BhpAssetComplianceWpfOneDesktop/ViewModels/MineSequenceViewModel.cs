using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
using System;
using System.Collections.Generic;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows.Media;
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Models.MineSequenceModels;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class MineSequenceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.MineSequence;
        protected override string MyPosterIcon { get; set; } = IconKeys.MineSequence;

        private string _myEscondidaImage;
        public string MyEscondidaImage
        {
            get { return _myEscondidaImage; }
            set { SetProperty(ref _myEscondidaImage, value); }
        }

        private string _myEscondidaNorteImage;
        public string MyEscondidaNorteImage
        {
            get { return _myEscondidaNorteImage; }
            set { SetProperty(ref _myEscondidaNorteImage, value); }
        }

        private ImageSource _myEscondidaImageSource;
        public ImageSource MyEscondidaImageSource
        {
            get { return _myEscondidaImageSource; }
            set { SetProperty(ref _myEscondidaImageSource, value); }
        }

        private ImageSource _myEscondidaNorteImageSource;
        public ImageSource MyEscondidaNorteImageSource
        {
            get { return _myEscondidaNorteImageSource; }
            set { SetProperty(ref _myEscondidaNorteImageSource, value); }
        }

        private bool _isEnabledLoadEscondidaImagePath;
        public bool IsEnabledLoadEscondidaImagePath
        {
            get { return _isEnabledLoadEscondidaImagePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaImagePath, value); }
        }

        private bool _isEnabledLoadEscondidaNorteImagePath;
        public bool IsEnabledLoadEscondidaNorteImagePath
        {
            get { return _isEnabledLoadEscondidaNorteImagePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaNorteImagePath, value); }
        }

        private bool _isEnabledGenerateTemplate;
        public bool IsEnabledGenerateTemplate
        {
            get { return _isEnabledGenerateTemplate; }
            set { SetProperty(ref _isEnabledGenerateTemplate, value); }
        }

        private bool _isEnabledLoadTemplateValues;
        public bool IsEnabledLoadTemplateValues
        {
            get { return _isEnabledLoadTemplateValues; }
            set { SetProperty(ref _isEnabledLoadTemplateValues, value); }
        }

        private string _myLastRefreshValues;
        public string MyLastRefreshValues
        {
            get { return _myLastRefreshValues; }
            set { SetProperty(ref _myLastRefreshValues, value); }
        }

        private string _myLastDateRefreshImages;
        public string MyLastDateRefreshImages
        {
            get { return _myLastDateRefreshImages; }
            set { SetProperty(ref _myLastDateRefreshImages, value); }
        }

        private DateTime _myDateActual;
        public DateTime MyDateActual
        {
            get { return _myDateActual; }
            set { SetProperty(ref _myDateActual, value); }
        }
     
        public DelegateCommand SelectEscondidaImageCommand { get; private set; }
        public DelegateCommand SelectEscondidaNorteImageCommand { get; private set; }
        public DelegateCommand LoadImagesCommand { get; private set; }
        public DelegateCommand GenerateMineSequenceTemplateCommand { get; private set; } 
        public DelegateCommand LoadMineSequenceTemplateCommand { get; private set; }
        public DelegateCommand ProcessAllMineSequenceValuesCommand { get; private set; }

        private readonly List<MineSequenceL1Expit> _l1Expit = new List<MineSequenceL1Expit>();

        private readonly List<MineSequenceAdherenceToB01L1> _adherenceToB01L1 = new List<MineSequenceAdherenceToB01L1>();
      
        private readonly List<MineSequenceDelayRecover> _delayRecover = new List<MineSequenceDelayRecover>();
        
        private readonly List<MineSequenceComments> _comments = new List<MineSequenceComments>();

        public MineSequenceViewModel()
        {
            IsEnabledLoadEscondidaImagePath = false;
            IsEnabledLoadEscondidaNorteImagePath = false;
            IsEnabledGenerateTemplate = false;
            MyDateActual = DateTime.Now;
            SelectEscondidaImageCommand = new DelegateCommand(EscondidaImagePath);
            SelectEscondidaNorteImageCommand = new DelegateCommand(EscondidaNorteImagePath);
            LoadImagesCommand = new DelegateCommand(LoadImages, CanProcess).ObservesProperty(() => IsEnabledLoadEscondidaImagePath).ObservesProperty(() => IsEnabledLoadEscondidaNorteImagePath); ;
            GenerateMineSequenceTemplateCommand = new DelegateCommand(GenerateMineSequenceTemplate);            
            LoadMineSequenceTemplateCommand = new DelegateCommand(LoadMineSequenceTemplate).ObservesCanExecute(() => IsEnabledGenerateTemplate);
        }

        private void EscondidaImagePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaImage = openFileDialog.FileName;
                MyEscondidaImageSource = new BitmapImage(new Uri(MyEscondidaImage));
            }
            IsEnabledLoadEscondidaImagePath = true;
        }

        private void EscondidaNorteImagePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaNorteImage = openFileDialog.FileName;
                MyEscondidaNorteImageSource = new BitmapImage(new Uri(MyEscondidaNorteImage));
            }
            IsEnabledLoadEscondidaNorteImagePath = true;
        }

        private bool CanProcess()
        {
            if (IsEnabledLoadEscondidaImagePath & IsEnabledLoadEscondidaNorteImagePath)
            {
                return true;
            }
            return false;
        }

        private void LoadImages()
        {
            var targetFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceCSVFilePath;
            var loadFileInfo = new FileInfo(targetFilePath);
            if (loadFileInfo.Exists)
            {
                if (targetFilePath.Substring(targetFilePath.Length - 27) == "MineSequencePictureData.csv")
                {
                    try
                    {
                        var openWriteCheck = File.OpenWrite(targetFilePath);
                        openWriteCheck.Close();

                        var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                        var findImageOnDate = ExportImageToCsv.SearchByDateMineSequence(newDate, targetFilePath);
                        if (findImageOnDate == -1)
                        {
                            ExportImageToCsv.AppendImageMineSequenceToCSV(MyEscondidaImage, MyEscondidaNorteImage, newDate, targetFilePath);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(targetFilePath, findImageOnDate);
                            ExportImageToCsv.AppendImageMineSequenceToCSV(MyEscondidaImage, MyEscondidaNorteImage, newDate, targetFilePath);
                        }
                        MyLastDateRefreshImages = $"{StringResources.Updated}: {DateTime.Now}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, StringResources.UploadError);
                    }
                }
                else
                {
                    var wrongFileMessage = $"{StringResources.WorksheetNotExist} {targetFilePath} {StringResources.IsTheRightOne}";
                    MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                }
            }
            else
            {
                var wrongFileMessage = $"{StringResources.WorksheetNotExist} {targetFilePath} {StringResources.ExistsOrNotSelect}";
                MessageBox.Show(wrongFileMessage, StringResources.UploadError);
            }         
        }

        private void GenerateMineSequenceTemplate()
        {
            var excelPackage = new ExcelPackage();

            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = MineSequenceConstants.MineSequenceExcelFileName;
            excelPackage.Workbook.Properties.Company = "BHP";

            var l1ExpitWorksheet = excelPackage.Workbook.Worksheets.Add(MineSequenceConstants.L1ExpitMineSequenceWorksheet);

            l1ExpitWorksheet.Cells["A1"].Value = "Expit Budget (t)";
            l1ExpitWorksheet.Cells["B1"].Value = "Expit Actual (%)";
            l1ExpitWorksheet.Cells["C1"].Value = "Budget Baseline";
            l1ExpitWorksheet.Cells["A1:C1"].Style.Font.Bold = true;
            l1ExpitWorksheet.Cells["A1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            l1ExpitWorksheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            l1ExpitWorksheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundExpitL1ExpitMineSequence));
            l1ExpitWorksheet.Cells["C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            l1ExpitWorksheet.Cells["C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundBudgetL1ExpitMineSequence));
            l1ExpitWorksheet.Column(3).Width = 18;

            for (var i = 1; i < 3; i++)
            {
                l1ExpitWorksheet.Cells[$"A{i}:C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                l1ExpitWorksheet.Cells[$"A{i}:C{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                l1ExpitWorksheet.Column(i).Width = 18;
            }

            var commentWorksheet = excelPackage.Workbook.Worksheets.Add(MineSequenceConstants.CommentsMineSequenceWorksheet);
            
            commentWorksheet.Cells["A1"].Value = "Tag";
            commentWorksheet.Cells["B1"].Value = "Comment";
            commentWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            commentWorksheet.Cells["A1:B1"].Style.Font.Bold = true;
            commentWorksheet.Column(1).Width = 15;
            commentWorksheet.Column(2).Width = 250;
            commentWorksheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            commentWorksheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundCommentsMineSequence));

            for (var i = 1; i < 21; i++)
            {
                commentWorksheet.Cells[$"A{i}:B{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                commentWorksheet.Cells[$"A{i}:B{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            var adherenceWorksheet = excelPackage.Workbook.Worksheets.Add(MineSequenceConstants.AdherenceMineSequenceWorksheet);
            
            var adherenceHeader = new List<string>() { "Unplanned Delay (t)", "Volume Ytd (%)", "Spatial Ytd (%)", "AdherenceL1 Ytd (%)" };

            for (var i = 0; i < adherenceHeader.Count; i++)
            {
                adherenceWorksheet.Cells[1, i + 1].Value = adherenceHeader[i];
                adherenceWorksheet.Cells[1, i + 1].Style.Font.Bold = true;
                adherenceWorksheet.Column(1 + i).Width = 21;
                adherenceWorksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                adherenceWorksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;          
            }
            adherenceWorksheet.Cells["B1:D1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.FontAdherenceMineSequence));
            adherenceWorksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundAdherenceMineSequence1));
            adherenceWorksheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundAdherenceMineSequence2));
            adherenceWorksheet.Cells["C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundAdherenceMineSequence3));
            adherenceWorksheet.Cells["D1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundAdherenceMineSequence4));

            for (var i = 1; i < 3; i++)
            {
                adherenceWorksheet.Cells[$"A{i}:D{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                adherenceWorksheet.Cells[$"A{i}:D{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            var delayrecoverWorksheet = excelPackage.Workbook.Worksheets.Add(MineSequenceConstants.DelayRecoverMineSequenceWorksheet);
           
            var delayrecoverHeader = new List<string>() { "Ytd PushBack (t)", "Phase Name", "DelayRecover Pushback (t)" };

            for (var i = 0; i < delayrecoverHeader.Count; i++)
            {
                delayrecoverWorksheet.Cells[1, i + 1].Value = delayrecoverHeader[i];
                delayrecoverWorksheet.Cells[1, i + 1].Style.Font.Bold = true;
                delayrecoverWorksheet.Column(1 + i).Width = 27;
                delayrecoverWorksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                delayrecoverWorksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                delayrecoverWorksheet.Cells[1, i + 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.FontDelayrecoverMineSequence));
                delayrecoverWorksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineSequenceTemplateColors.HeaderBackgroundDelayrecoverMineSequence));
            }

            delayrecoverWorksheet.Cells[$"A1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            delayrecoverWorksheet.Cells[$"A2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            for (var i = 1; i < 11; i++)
            {
                delayrecoverWorksheet.Cells[$"B{i}:C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                delayrecoverWorksheet.Cells[$"A{i}:C{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = MineSequenceConstants.MineSequenceExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                File.WriteAllBytes(dialog.FileName, fileText);
            }
            IsEnabledGenerateTemplate = true;
        }

        private void LoadMineSequenceTemplate()
        {
            _l1Expit.Clear();
            _adherenceToB01L1.Clear();
            _delayRecover.Clear();
            _comments.Clear();

            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var l1ExpitTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.L1ExpitMineSequenceWorksheet];
                var adherenceTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.AdherenceMineSequenceWorksheet];
                var delayrecoverTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.DelayRecoverMineSequenceWorksheet];
                var commentTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.CommentsMineSequenceWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - MineSequenceConstants.MineSequenceExcelFileName.Length) == MineSequenceConstants.MineSequenceExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var openWriteCheck = File.OpenWrite(openFileDialog.FileName);
                        openWriteCheck.Close();

                        if (l1ExpitTemplateWorksheet.Cells[2, 1].Value == null)
                            l1ExpitTemplateWorksheet.Cells[2, 1].Value = -99;
                        if (l1ExpitTemplateWorksheet.Cells[2, 2].Value == null)
                            l1ExpitTemplateWorksheet.Cells[2, 2].Value = -9900;
                        if (l1ExpitTemplateWorksheet.Cells[2, 3].Value == null)
                            l1ExpitTemplateWorksheet.Cells[2, 3].Value = " ";

                        _l1Expit.Add(new MineSequenceL1Expit()
                        {
                            ExpitBudgetTonnes = double.Parse(l1ExpitTemplateWorksheet.Cells[2, 1].Value.ToString()),
                            ExpitActualPercent = double.Parse(l1ExpitTemplateWorksheet.Cells[2, 2].Value.ToString())/100,
                            BudgetBaseline = l1ExpitTemplateWorksheet.Cells[2, 3].Value.ToString()
                        });

                        if (adherenceTemplateWorksheet.Cells[2, 1].Value == null)
                            adherenceTemplateWorksheet.Cells[2, 1].Value = -99;
                        if (adherenceTemplateWorksheet.Cells[2, 2].Value == null)
                            adherenceTemplateWorksheet.Cells[2, 2].Value = -9900;
                        if (adherenceTemplateWorksheet.Cells[2, 3].Value == null)
                            adherenceTemplateWorksheet.Cells[2, 3].Value = -9900;
                        if (adherenceTemplateWorksheet.Cells[2, 4].Value == null)
                            adherenceTemplateWorksheet.Cells[2, 4].Value = -9900;

                        _adherenceToB01L1.Add(new MineSequenceAdherenceToB01L1()
                        {
                            UnplannedDelayTonnes = double.Parse(adherenceTemplateWorksheet.Cells[2, 1].Value.ToString()),
                            VolumeYtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 2].Value.ToString())/100,
                            SpatialYtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 3].Value.ToString())/100,
                            AdherenceL1YtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 4].Value.ToString())/100
                        });

                        var rows = delayrecoverTemplateWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows; i++)
                        {
                            if (delayrecoverTemplateWorksheet.Cells[i + 1, 2].Value != null)
                            {
                                if (delayrecoverTemplateWorksheet.Cells[2, 1].Value == null)
                                    delayrecoverTemplateWorksheet.Cells[2, 1].Value = -99;
                                if (delayrecoverTemplateWorksheet.Cells[i + 1, 3].Value == null)
                                    delayrecoverTemplateWorksheet.Cells[i + 1, 3].Value = -99;

                                _delayRecover.Add(new MineSequenceDelayRecover()
                                {
                                    YtdPushBackTonnes = double.Parse(delayrecoverTemplateWorksheet.Cells[2, 1].Value.ToString()),
                                    PhaseName = delayrecoverTemplateWorksheet.Cells[i + 1, 2].Value.ToString(),
                                    DelayRecoverPushbackTonnes = double.Parse(delayrecoverTemplateWorksheet.Cells[i + 1, 3].Value.ToString()),
                                });
                            }
                        }

                        var rows2 = commentTemplateWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows2; i++)
                        {
                            if (commentTemplateWorksheet.Cells[i + 1, 2].Value != null)
                            {
                                string tag;

                                if (commentTemplateWorksheet.Cells[i + 1, 1].Value.ToNullSafeString() == "")
                                {
                                    tag = "All";
                                }
                                else
                                {
                                    tag = commentTemplateWorksheet.Cells[i + 1, 1].Value.ToString();
                                }

                                _comments.Add(new MineSequenceComments()
                                {
                                    Tag = tag,
                                    Comment = commentTemplateWorksheet.Cells[i + 1, 2].Value.ToString()
                                });
                            }
                        }
                        excelPackage.Dispose();

                        var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineSequenceExcelFilePath;
                        var loadFileInfo = new FileInfo(loadFilePath);

                        if (loadFileInfo.Exists)
                        {
                            var package = new ExcelPackage(loadFileInfo);
                            var l1ExpitWorksheet = package.Workbook.Worksheets[MineSequenceConstants.L1ExpitMineSequenceWorksheet];
                            var adherenceWorksheet = package.Workbook.Worksheets[MineSequenceConstants.AdherenceMineSequenceWorksheet];
                            var delayrecoverWorksheet = package.Workbook.Worksheets[MineSequenceConstants.DelayRecoverMineSequenceWorksheet];
                            var commentWorksheet = package.Workbook.Worksheets[MineSequenceConstants.CommentsMineSequenceWorksheet];

                            if (l1ExpitWorksheet != null & adherenceWorksheet != null & delayrecoverWorksheet != null & commentWorksheet != null)
                            {
                                var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                                var lastRow1 = l1ExpitWorksheet.Dimension.End.Row + 1;

                                l1ExpitWorksheet.Cells[lastRow1, 1].Value = newDate;
                                l1ExpitWorksheet.Cells[lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                l1ExpitWorksheet.Cells[lastRow1, 2].Value = _l1Expit[0].ExpitBudgetTonnes;
                                l1ExpitWorksheet.Cells[lastRow1, 3].Value = _l1Expit[0].ExpitActualPercent;
                                l1ExpitWorksheet.Cells[lastRow1, 4].Value = _l1Expit[0].BudgetBaseline;

                                var lastRow2 = adherenceWorksheet.Dimension.End.Row + 1;

                                adherenceWorksheet.Cells[lastRow2, 1].Value = newDate;
                                adherenceWorksheet.Cells[lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                adherenceWorksheet.Cells[lastRow2, 2].Value = _adherenceToB01L1[0].UnplannedDelayTonnes;
                                adherenceWorksheet.Cells[lastRow2, 3].Value = _adherenceToB01L1[0].VolumeYtdPercent;
                                adherenceWorksheet.Cells[lastRow2, 4].Value = _adherenceToB01L1[0].SpatialYtdPercent;
                                adherenceWorksheet.Cells[lastRow2, 5].Value = _adherenceToB01L1[0].AdherenceL1YtdPercent;

                                var lastRow3 = delayrecoverWorksheet.Dimension.End.Row + 1;

                                for (var i = 0; i < _delayRecover.Count; i++)
                                {
                                    delayrecoverWorksheet.Cells[i + lastRow3, 1].Value = newDate;
                                    delayrecoverWorksheet.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    delayrecoverWorksheet.Cells[i + lastRow3, 2].Value = _delayRecover[i].YtdPushBackTonnes;
                                    delayrecoverWorksheet.Cells[i + lastRow3, 3].Value = _delayRecover[i].PhaseName;
                                    delayrecoverWorksheet.Cells[i + lastRow3, 4].Value = _delayRecover[i].DelayRecoverPushbackTonnes;
                                }

                                var lastRow4 = commentWorksheet.Dimension.End.Row + 1;

                                for (var i = 0; i < _comments.Count; i++)
                                {
                                    commentWorksheet.Cells[i + lastRow4, 1].Value = newDate;
                                    commentWorksheet.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    commentWorksheet.Cells[i + lastRow4, 2].Value = _comments[i].Tag;
                                    commentWorksheet.Cells[i + lastRow4, 3].Value = _comments[i].Comment;
                                }
                                byte[] fileText = package.GetAsByteArray();
                                File.WriteAllBytes(loadFilePath, fileText);
                                MyLastRefreshValues = $"{StringResources.Updated}: {DateTime.Now}";
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
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, StringResources.UploadError);
                    }
                }
                else
                {
                    var wrongFileMessage = $"{StringResources.WrongUploadedFile} {openFilePath.FullName} {StringResources.IsTheRightOne}";
                    MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                }
            }
        }      
    }
}
