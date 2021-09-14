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
using System.Drawing.Printing;
using System.Windows.Media;
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;
using BhpAssetComplianceWpfOneDesktop.Models.MineSequenceModels;
using BhpAssetComplianceWpfOneDesktop.Models.BlastingInventoryModels;
using OfficeOpenXml.DataValidation;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class BlastingInventoryViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.BlastingInventory;
        protected override string MyPosterIcon { get; set; } = IconKeys.Blasting;

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

        private DateTime _myStartDateRange;
        public DateTime MyStartDateRange
        {
            get { return _myStartDateRange; }
            set { SetProperty(ref _myStartDateRange, value); }
        }

        private DateTime _myEndDateRange;
        public DateTime MyEndDateRange
        {
            get { return _myEndDateRange; }
            set { SetProperty(ref _myEndDateRange, value); }
        }

        private IEnumerable<DateTime> _myDateRange;
        public IEnumerable<DateTime> MyDateRange
        {
            get { return _myDateRange; }
            set { SetProperty(ref _myDateRange, value); }
        }

        public DelegateCommand SelectEscondidaImageCommand { get; private set; }
        public DelegateCommand SelectEscondidaNorteImageCommand { get; private set; }
        public DelegateCommand LoadImagesCommand { get; private set; }
        public DelegateCommand GenerateBlastingInventoryTemplateCommand { get; private set; }
        public DelegateCommand LoadBlastingInventoryTemplateCommand { get; private set; }


        private readonly List<BlastingInventoryBlast> _blast = new List<BlastingInventoryBlast>();

        private readonly List<BlastingInventoryPhaseBlast> _phaseBlast = new List<BlastingInventoryPhaseBlast>();

        private readonly List<BlastingInventoryShovels> _shovels = new List<BlastingInventoryShovels>();

        private readonly List<BlastingInventoryWeeklySummary> _weeklySummary = new List<BlastingInventoryWeeklySummary>();



        private readonly List<MineSequenceL1Expit> _l1Expit = new List<MineSequenceL1Expit>();

        private readonly List<MineSequenceAdherenceToB01L1> _adherenceToB01L1 = new List<MineSequenceAdherenceToB01L1>();

        private readonly List<MineSequenceDelayRecover> _delayRecover = new List<MineSequenceDelayRecover>();

        private readonly List<MineSequenceComments> _comments = new List<MineSequenceComments>();

        public BlastingInventoryViewModel()
        {
            IsEnabledLoadEscondidaImagePath = false;
            IsEnabledLoadEscondidaNorteImagePath = false;
            IsEnabledGenerateTemplate = false;
            MyDateActual = DateTime.Now;
            MyStartDateRange = DateTime.Now;
            MyEndDateRange = DateTime.Now;
            SelectEscondidaImageCommand = new DelegateCommand(EscondidaImagePath);
            SelectEscondidaNorteImageCommand = new DelegateCommand(EscondidaNorteImagePath);
            LoadImagesCommand = new DelegateCommand(LoadImages, CanProcess).ObservesProperty(() => IsEnabledLoadEscondidaImagePath).ObservesProperty(() => IsEnabledLoadEscondidaNorteImagePath); ;
            GenerateBlastingInventoryTemplateCommand = new DelegateCommand(GenerateBlastingInventoryTemplate);
            LoadBlastingInventoryTemplateCommand = new DelegateCommand(LoadBlastingInventoryTemplate).ObservesCanExecute(() => IsEnabledGenerateTemplate);
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
            var targetFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.BlastingInventoryCSVFilePath;
            var loadFileInfo = new FileInfo(targetFilePath);
            if (loadFileInfo.Exists)
            {
                if (targetFilePath.Substring(targetFilePath.Length - 32) == "BlastingInventoryPictureData.csv")
                {
                    try
                    {
                        var openWriteCheck = File.OpenWrite(targetFilePath);
                        openWriteCheck.Close();

   
                        var findImageOnDate = ExportImageToCsv.SearchByDateMineSequence(MyDateActual, targetFilePath);
                        if (findImageOnDate == -1)
                        {
                            ExportImageToCsv.AppendImageMineSequenceToCSV(MyEscondidaImage, MyEscondidaNorteImage, MyDateActual, targetFilePath);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(targetFilePath, findImageOnDate);
                            ExportImageToCsv.AppendImageMineSequenceToCSV(MyEscondidaImage, MyEscondidaNorteImage, MyDateActual, targetFilePath);
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

        private void GenerateBlastingInventoryTemplate()
        {
            var excelPackage = new ExcelPackage();

            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = BlastingInventoryConstants.BlastingInventoryWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var blastingEscondidaWorksheet = excelPackage.Workbook.Worksheets.Add(BlastingInventoryConstants.BlastingEscondidaBlastingInventoryWorksheet);
            var blastingEscondidaNorteWorksheet = excelPackage.Workbook.Worksheets.Add(BlastingInventoryConstants.BlastingEscondidaNorteBlastingInventoryWorksheet);
            var shovelsWorksheet = excelPackage.Workbook.Worksheets.Add(BlastingInventoryConstants.ShovelsBlastingInventoryWorksheet);
            var weeklySummaryWorksheet = excelPackage.Workbook.Worksheets.Add(BlastingInventoryConstants.WeeklySummaryBlastingInventoryWorksheet);


            MyDateRange = TemplateDates.GetDateRange(MyStartDateRange, MyEndDateRange);

            blastingEscondidaWorksheet.Cells["A1:D1"].Merge = true;
            blastingEscondidaWorksheet.Cells["A1"].Value = "Blasting Inventory - Escondida";
            blastingEscondidaWorksheet.Cells["A1"].Style.Font.Bold = true;
            blastingEscondidaWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            blastingEscondidaWorksheet.Cells["A1"].Style.Font.Size = 12;

            blastingEscondidaWorksheet.Cells["H2:Q2"].Merge = true;
            blastingEscondidaWorksheet.Cells["H2"].Value = "Phase";
            blastingEscondidaWorksheet.Cells["H2"].Style.Font.Bold = true;
            blastingEscondidaWorksheet.Cells["H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            blastingEscondidaWorksheet.Cells["H2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            blastingEscondidaWorksheet.Cells["H2"].Style.Fill.BackgroundColor.SetColor(TemplateColors.OrangeBackground);
            blastingEscondidaWorksheet.Cells["H2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            blastingEscondidaWorksheet.Cells["H2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            blastingEscondidaWorksheet.Cells["H2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            var escondidaBlastingHeader = new List<string>()
            {
                "Month", "Day", "Sulphide (ton)", "Others (ton)", "Day Blast (ton)", "Events/Week", "Target", "S2B", "S3C",
                "MP1", "PL1", "N15", "N16", "N17", "E04", " E06", "E07"
            };

            for (var i = 0; i < escondidaBlastingHeader.Count; i++)
            {
                blastingEscondidaWorksheet.Cells[3, i + 1].Value = escondidaBlastingHeader[i];
                blastingEscondidaWorksheet.Cells[3, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                blastingEscondidaWorksheet.Cells[3, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                blastingEscondidaWorksheet.Column(i + 1).Width = 12;
            }

            blastingEscondidaWorksheet.Cells["A3:Q3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            blastingEscondidaWorksheet.Cells["A3:Q3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            blastingEscondidaWorksheet.Cells["A3:Q3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            blastingEscondidaWorksheet.Cells["A3:G3"].Style.Font.Bold = true;
            blastingEscondidaWorksheet.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.BlueBackground);
            blastingEscondidaWorksheet.Cells["C3:E3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightGreenBackground);
            blastingEscondidaWorksheet.Cells["H3:Q3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightOrangeBackground);
            blastingEscondidaWorksheet.Cells["F3:G3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightGrayBackground);
            blastingEscondidaWorksheet.Cells["A3:B3"].Style.Font.Color.SetColor(TemplateColors.WhiteFont);

            var aux = 1;
            foreach (DateTime value in MyDateRange)
            {
                blastingEscondidaWorksheet.Cells[3 + aux, 1].Value = value.Month;
                blastingEscondidaWorksheet.Cells[3 + aux, 2].Value = value.Day;
                blastingEscondidaWorksheet.Cells[$"A{3 + aux}:Q{3 + aux}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                blastingEscondidaWorksheet.Cells[$"A{3 + aux}:Q{3 + aux}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                aux++;
            }

            var escondidaNorteBlastingHeader = new List<string>()
            {
                "Month", "Day", "Sulphide (ton)", "Others (ton)", "Day Blast (ton)", "Events/Week", "Target",
                "N03", " N04", " N568", "N05", "N06", "N07", "N08", "N9", "N10", " N11"
            };

            blastingEscondidaNorteWorksheet.Cells["A1:D1"].Merge = true;
            blastingEscondidaNorteWorksheet.Cells["A1"].Value = "Blasting Inventory - Escondida Norte";
            blastingEscondidaNorteWorksheet.Cells["A1"].Style.Font.Bold = true;
            blastingEscondidaNorteWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            blastingEscondidaNorteWorksheet.Cells["A1"].Style.Font.Size = 12;

            blastingEscondidaNorteWorksheet.Cells["H2:Q2"].Merge = true;
            blastingEscondidaNorteWorksheet.Cells["H2"].Value = "Phase";
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Font.Bold = true;
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Fill.BackgroundColor.SetColor(TemplateColors.OrangeBackground);
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            blastingEscondidaNorteWorksheet.Cells["H2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < escondidaNorteBlastingHeader.Count; i++)
            {
                blastingEscondidaNorteWorksheet.Cells[3, i + 1].Value = escondidaNorteBlastingHeader[i];
                blastingEscondidaNorteWorksheet.Cells[3, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                blastingEscondidaNorteWorksheet.Cells[3, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                blastingEscondidaNorteWorksheet.Column(i + 1).Width = 12;
            }

            blastingEscondidaNorteWorksheet.Cells["A3:Q3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            blastingEscondidaNorteWorksheet.Cells["A3:Q3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            blastingEscondidaNorteWorksheet.Cells["A3:Q3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            blastingEscondidaNorteWorksheet.Cells["A3:G3"].Style.Font.Bold = true;
            blastingEscondidaNorteWorksheet.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.BlueBackground);
            blastingEscondidaNorteWorksheet.Cells["C3:E3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightGreenBackground);
            blastingEscondidaNorteWorksheet.Cells["H3:Q3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightOrangeBackground);
            blastingEscondidaNorteWorksheet.Cells["F3:G3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightGrayBackground);
            blastingEscondidaNorteWorksheet.Cells["A3:B3"].Style.Font.Color.SetColor(TemplateColors.WhiteFont);

            var aux2 = 1;
            foreach (DateTime value in MyDateRange)
            {
                blastingEscondidaNorteWorksheet.Cells[3 + aux2, 1].Value = value.Month;
                blastingEscondidaNorteWorksheet.Cells[3 + aux2, 2].Value = value.Day;
                blastingEscondidaNorteWorksheet.Cells[$"A{3 + aux2}:Q{3 + aux2}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                blastingEscondidaNorteWorksheet.Cells[$"A{3 + aux2}:Q{3 + aux2}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                aux2++;
            }

            shovelsWorksheet.Cells["A1:D1"].Merge = true;
            shovelsWorksheet.Cells["A1"].Value = "Shovel Distribution - Escondida";
            shovelsWorksheet.Cells["A1"].Style.Font.Bold = true;
            shovelsWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            shovelsWorksheet.Cells["A1"].Style.Font.Size = 12;

            shovelsWorksheet.Cells["H1:L1"].Merge = true;
            shovelsWorksheet.Cells["H1"].Value = "Shovel Distribution - Escondida Norte";
            shovelsWorksheet.Cells["H1"].Style.Font.Bold = true;
            shovelsWorksheet.Cells["H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            shovelsWorksheet.Cells["H1"].Style.Font.Size = 12;

            var shovelsHeader = new List<string>()
            {
                "Month", "Day", "Type", "Phase", "Value"
            };



            for (var i = 0; i < shovelsHeader.Count; i++)
            {
                shovelsWorksheet.Cells[3, i + 1].Value = shovelsHeader[i];
                shovelsWorksheet.Cells[3, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                shovelsWorksheet.Cells[3, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                shovelsWorksheet.Column(i + 1).Width = 8;

                shovelsWorksheet.Cells[3, i + 8].Value = shovelsHeader[i];
                shovelsWorksheet.Cells[3, i + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                shovelsWorksheet.Cells[3, i + 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                shovelsWorksheet.Column(i + 6).Width = 8;
            }

            shovelsWorksheet.Cells["A3:L3"].Style.Font.Bold = true;
            shovelsWorksheet.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.BlueBackground);
            shovelsWorksheet.Cells["A3:B3"].Style.Font.Color.SetColor(TemplateColors.WhiteFont);
            shovelsWorksheet.Cells["H3:I3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.BlueBackground);
            shovelsWorksheet.Cells["H3:I3"].Style.Font.Color.SetColor(TemplateColors.WhiteFont);

            shovelsWorksheet.Cells["C3:E3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightGreenBackground);
            shovelsWorksheet.Cells["J3:L3"].Style.Fill.BackgroundColor.SetColor(TemplateColors.LightOrangeBackground);

            shovelsWorksheet.Cells["A3:E3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            shovelsWorksheet.Cells["H3:L3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            shovelsWorksheet.Cells["A3:E3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            shovelsWorksheet.Cells["H3:L3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            shovelsWorksheet.Cells["A4:B4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            shovelsWorksheet.Cells["H4:I4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;

            shovelsWorksheet.Cells["A3:E3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            shovelsWorksheet.Cells["H3:L3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            shovelsWorksheet.Cells["A4:E4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            shovelsWorksheet.Cells["H4:L4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < 30; i++)
            {
                shovelsWorksheet.Cells[$"C{4 + i}:E{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                shovelsWorksheet.Cells[i + 4, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                shovelsWorksheet.Cells[$"C{4+i}:E{4+i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                var noteTypeListDataValidation = shovelsWorksheet.Cells[i + 4, 3].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteTypeListDataValidation.AllowBlank = false;
                noteTypeListDataValidation.Formula.Values.Add("Lastre");
                noteTypeListDataValidation.Formula.Values.Add("Mineral");
                noteTypeListDataValidation.ShowErrorMessage = true;
                noteTypeListDataValidation.Error = "Select from List of Values ...";

                shovelsWorksheet.Cells[$"J{4 + i}:L{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                shovelsWorksheet.Cells[i + 4, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                shovelsWorksheet.Cells[$"J{4 + i}:L{4 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                var noteStateListDataValidations = shovelsWorksheet.Cells[i + 4, 10].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteStateListDataValidations.AllowBlank = false;
                noteStateListDataValidations.Formula.Values.Add("Lastre");
                noteStateListDataValidations.Formula.Values.Add("Mineral");
                noteStateListDataValidations.ShowErrorMessage = true;
                noteStateListDataValidations.Error = "Select from List of Values ...";
            }


            weeklySummaryWorksheet.Cells["A1:D1"].Merge = true;
            weeklySummaryWorksheet.Cells["A1"].Value = "Weekly Summary";
            weeklySummaryWorksheet.Cells["A1"].Style.Font.Bold = true;
            weeklySummaryWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            weeklySummaryWorksheet.Cells["A1"].Style.Font.Size = 12;

            var weeklySummaryHeader = new List<string>()
            {
                "Week","Avg ton", "Sum Events","Target ","Avg ton", "Sum Events","Target "

            };

            weeklySummaryWorksheet.Cells["B3:D3"].Merge = true;
            weeklySummaryWorksheet.Cells["B3"].Value = "Escondida";
            weeklySummaryWorksheet.Cells["B3:G3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            weeklySummaryWorksheet.Cells["B3:G3"].Style.Font.Bold = true;
            weeklySummaryWorksheet.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            weeklySummaryWorksheet.Cells["E3"].Value = "Escondida Norte";

            weeklySummaryWorksheet.Cells["B3:G3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            weeklySummaryWorksheet.Cells[3, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            weeklySummaryWorksheet.Cells["B3:G3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


            weeklySummaryWorksheet.Cells[3,2].Style.Fill.BackgroundColor.SetColor(TemplateColors.GreenBackground);
            weeklySummaryWorksheet.Cells[3,5].Style.Fill.BackgroundColor.SetColor(TemplateColors.OrangeBackground);

            for (var i = 0; i < weeklySummaryHeader.Count; i++)
            {
                weeklySummaryWorksheet.Cells[4, i + 1].Value = weeklySummaryHeader[i];
                weeklySummaryWorksheet.Cells[4, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                weeklySummaryWorksheet.Cells[4, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                weeklySummaryWorksheet.Column(i + 1).Width = 10;
                weeklySummaryWorksheet.Cells[4, i+1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                weeklySummaryWorksheet.Cells[4, i+1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                weeklySummaryWorksheet.Cells[4, i+1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            for (var i = 0; i < 2; i++)
            {

            }
            weeklySummaryWorksheet.Cells["B3:G3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            weeklySummaryWorksheet.Cells[3,2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            weeklySummaryWorksheet.Cells["B3:G3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            weeklySummaryWorksheet.Cells[4,1].Style.Font.Bold = true;

            for (var i = 0; i < 4; i++)
            {
                weeklySummaryWorksheet.Cells[3, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                weeklySummaryWorksheet.Cells[$"A{5+i}:G{5+i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            
            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = BlastingInventoryConstants.BlastingInventoryExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                File.WriteAllBytes(dialog.FileName, fileText);
            }
            IsEnabledGenerateTemplate = true;
        }

        private void LoadBlastingInventoryTemplate()
        {
            _blast.Clear();
            _phaseBlast.Clear();
            _shovels.Clear();
            _weeklySummary.Clear(); 

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
                var blastingEscondidaTemplateWorksheet = excelPackage.Workbook.Worksheets[BlastingInventoryConstants.BlastingEscondidaBlastingInventoryWorksheet];
                var blastingEscondidaNorteTemplateWorksheet = excelPackage.Workbook.Worksheets[BlastingInventoryConstants.BlastingEscondidaNorteBlastingInventoryWorksheet];
                var shovelsTemplateWorksheet = excelPackage.Workbook.Worksheets[BlastingInventoryConstants.ShovelsBlastingInventoryWorksheet];
                var weeklySummaryTemplateWorksheet = excelPackage.Workbook.Worksheets[BlastingInventoryConstants.WeeklySummaryBlastingInventoryWorksheet];


                

                var l1ExpitTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.L1ExpitMineSequenceWorksheet];
                var adherenceTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.AdherenceMineSequenceWorksheet];
                var delayrecoverTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.DelayRecoverMineSequenceWorksheet];
                var commentTemplateWorksheet = excelPackage.Workbook.Worksheets[MineSequenceConstants.CommentsMineSequenceWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - BlastingInventoryConstants.BlastingInventoryExcelFileName.Length) == BlastingInventoryConstants.BlastingInventoryExcelFileName)
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
                            ExpitActualPercent = double.Parse(l1ExpitTemplateWorksheet.Cells[2, 2].Value.ToString()) / 100,
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
                            VolumeYtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 2].Value.ToString()) / 100,
                            SpatialYtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 3].Value.ToString()) / 100,
                            AdherenceL1YtdPercent = double.Parse(adherenceTemplateWorksheet.Cells[2, 4].Value.ToString()) / 100
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