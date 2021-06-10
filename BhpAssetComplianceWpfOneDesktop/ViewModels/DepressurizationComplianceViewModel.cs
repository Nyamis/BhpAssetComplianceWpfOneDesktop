using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
using System;
using System.Collections.Generic;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Style;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using BhpAssetComplianceWpfOneDesktop.Constants;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Models.DepressurizationComplianceModels;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class DepressurizationComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.DepressurizationCompliance;
        protected override string MyPosterIcon { get; set; } = IconKeys.Depressurization;

        private string _myImage;
        public string MyImage
        {
            get { return _myImage; }
            set { SetProperty(ref _myImage, value); }
        }

        private ImageSource _myImageSource;
        public ImageSource MyImageSource
        {
            get { return _myImageSource; }
            set { SetProperty(ref _myImageSource, value); }
        }

        private string _myLastDateRefreshMonthlyImage;
        public string MyLastDateRefreshMonthlyImage
        {
            get { return _myLastDateRefreshMonthlyImage; }
            set { SetProperty(ref _myLastDateRefreshMonthlyImage, value); }
        }

        private string _myLastDateRefreshMonthlyValues;
        public string MyLastDateRefreshMonthlyValues
        {
            get { return _myLastDateRefreshMonthlyValues; }
            set { SetProperty(ref _myLastDateRefreshMonthlyValues, value); }
        }

        private string _myLastDateRefreshTargetValues;
        public string MyLastDateRefreshTargetValues
        {
            get { return _myLastDateRefreshTargetValues; }
            set { SetProperty(ref _myLastDateRefreshTargetValues, value); }
        }

        private DateTime _myDateActual;
        public DateTime MyDateActual
        {
            get { return _myDateActual; }
            set { SetProperty(ref _myDateActual, value); }
        }

        private int _myFiscalYear;
        public int MyFiscalYear
        {
            get { return _myFiscalYear; }
            set { SetProperty(ref _myFiscalYear, value); }
        }

        private bool _isEnabledLoadImagePath;
        public bool IsEnabledLoadImagePath
        {
            get { return _isEnabledLoadImagePath; }
            set { SetProperty(ref _isEnabledLoadImagePath, value); }
        }

        private bool _isEnabledGenerateMonthlyTemplate;
        public bool IsEnabledGenerateMonthlyTemplate
        {
            get { return _isEnabledGenerateMonthlyTemplate; }
            set { SetProperty(ref _isEnabledGenerateMonthlyTemplate, value); }
        }

        private bool _isEnabledGenerateTargetTemplate;
        public bool IsEnabledGenerateTargetTemplate
        {
            get { return _isEnabledGenerateTargetTemplate; }
            set { SetProperty(ref _isEnabledGenerateTargetTemplate, value); }
        }
        public DelegateCommand SelectImageCommand { get; private set; }
        public DelegateCommand LoadImageCommand { get; private set; }
        public DelegateCommand GenerateMonthlyDepressurizationTemplateCommand { get; private set; }
        public DelegateCommand LoadMonthlyDepressurizationTemplateCommand { get; private set; }
        public DelegateCommand GenerateTargetDepressurizationTemplateCommand { get; private set; }
        public DelegateCommand LoadTargetDepressurizationTemplateCommand { get; private set; }

        private readonly List<DepressurizationComplianceMonthlyCompliance> _monthlyCompliance = new List<DepressurizationComplianceMonthlyCompliance>();      
        private readonly List<DepressurizationComplianceTargetCompliance> _targetCompliance = new List<DepressurizationComplianceTargetCompliance>();

        public DepressurizationComplianceViewModel()
        {
            MyDateActual = DateTime.Now;
            MyFiscalYear = MyDateActual.Year;
            IsEnabledLoadImagePath = false;
            IsEnabledGenerateMonthlyTemplate = false;
            IsEnabledGenerateTargetTemplate = false;
            SelectImageCommand = new DelegateCommand(SelectImagePath);
            LoadImageCommand = new DelegateCommand(LoadImage).ObservesCanExecute(() => IsEnabledLoadImagePath);
            GenerateMonthlyDepressurizationTemplateCommand = new DelegateCommand(GenerateDepressurizationMonthlyTemplate);
            LoadMonthlyDepressurizationTemplateCommand = new DelegateCommand(LoadDepressurizationMonthlyTemplate).ObservesCanExecute(() => IsEnabledGenerateMonthlyTemplate);
            GenerateTargetDepressurizationTemplateCommand = new DelegateCommand(GenerateDepressurizationTargetTemplate);
            LoadTargetDepressurizationTemplateCommand = new DelegateCommand(LoadDepressurizationTargetTemplate).ObservesCanExecute(() => IsEnabledGenerateTargetTemplate);            
        }

        private void SelectImagePath()
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
                MyImage = openFileDialog.FileName;
                MyImageSource = new BitmapImage(new Uri(openFileDialog.FileName));
            }
            IsEnabledLoadImagePath = true;
        }

        private void LoadImage()
        {
            var targetFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceCSVFilePath;
            var loadFileInfo = new FileInfo(targetFilePath);
            if (loadFileInfo.Exists)
            {
                if (targetFilePath.Substring(targetFilePath.Length - 41) == "DepressurizationCompliancePictureData.csv")
                {
                    try
                    {
                        var openWriteCheck = File.OpenWrite(targetFilePath);
                        openWriteCheck.Close();

                        var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                        var findImageOnDate = ExportImageToCsv.SearchByDateMineSequence(newDate, targetFilePath);
                        if (findImageOnDate == -1)
                        {
                            ExportImageToCsv.AppendImageDepressurizationToCSV(MyImage, newDate, targetFilePath);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(targetFilePath, findImageOnDate);
                            ExportImageToCsv.AppendImageDepressurizationToCSV(MyImage, newDate, targetFilePath);
                        }
                        MyLastDateRefreshMonthlyImage = $"{StringResources.Updated}: {DateTime.Now}";
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

        private void GenerateDepressurizationMonthlyTemplate()
        {
            var headers = new List<string> { "Wall", "Observado (kPa)", "Compliance (%)", "Pit" };
            var zones = new List<string> { "Pared Noreste Fuera Rajo", "Pared Noreste", "Pared Noreste Talud Bajo", "Pared Los Colorados", "Pared Los Colorados Talud Bajo", "Pared Este Fuera Rajo", "Pared Este Talud Medio" };            
            var excelPackage = new ExcelPackage();

            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = DepressurizationComplianceConstants.MonthlyDepressurizationWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";
            var worksheet = excelPackage.Workbook.Worksheets.Add(DepressurizationComplianceConstants.MonthlyDepressurizationWorksheet);

            for (var i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(1 + i).Width = 16;
                worksheet.Cells[1, 1 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, 1 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(DepressurizationComplianceTemplateColors.HeaderBackgroundMonthlyDepressurizationCompliance));
            }
            worksheet.Column(1).Width = 27;

            for (var i = 0; i < zones.Count; i++)
            {
                worksheet.Cells[2 + i, 1].Value = zones[i];
            }

            for (var i = 0; i <= 12; i++)
            {
                worksheet.Cells[i + 1, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = DepressurizationComplianceConstants.MonthlyDepressurizationExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledGenerateMonthlyTemplate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadDepressurizationMonthlyTemplate()
        {
            _monthlyCompliance.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var templateWorksheet = excelPackage.Workbook.Worksheets[DepressurizationComplianceConstants.MonthlyDepressurizationWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - DepressurizationComplianceConstants.MonthlyDepressurizationExcelFileName.Length) == DepressurizationComplianceConstants.MonthlyDepressurizationExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var openWriteCheck = File.OpenWrite(openFileDialog.FileName);
                        openWriteCheck.Close();

                        var rows = templateWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows; i++)
                        {
                            if (templateWorksheet.Cells[i + 1, 1].Value != null)
                            {
                               if (templateWorksheet.Cells[1 + i, 2].Value == null)
                                    templateWorksheet.Cells[1 + i, 2].Value = -99000;

                                if (templateWorksheet.Cells[1 + i, 3].Value == null)
                                    templateWorksheet.Cells[1 + i, 3].Value = -9900;


                                if (templateWorksheet.Cells[1 + i, 4].Value == null)
                                    templateWorksheet.Cells[1 + i, 4].Value = " ";

                                _monthlyCompliance.Add(new DepressurizationComplianceMonthlyCompliance()
                                {
                                    Zone = templateWorksheet.Cells[1 + i, 1].Value.ToString(),
                                    Observado = double.Parse(templateWorksheet.Cells[1 + i, 2].Value.ToString())/1000,
                                    Compliance = double.Parse(templateWorksheet.Cells[1 + i, 3].Value.ToString())/100,
                                    Pit = templateWorksheet.Cells[1 + i, 4].Value.ToString()
                                });
                            }
                        }
                        excelPackage.Dispose();

                        var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceExcelFilePath;
                        var loadFileInfo = new FileInfo(loadFilePath);

                        if (loadFileInfo.Exists)
                        {
                            var package = new ExcelPackage(loadFileInfo);
                            var worksheet = package.Workbook.Worksheets[DepressurizationComplianceConstants.MonthlyDepressurizationSpotfireWorksheet];

                            if (worksheet != null)
                            {
                                var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                                var lastRow = worksheet.Dimension.End.Row + 1;

                                for (var i = 0; i < _monthlyCompliance.Count; i++)
                                {
                                    worksheet.Cells[i + lastRow, 1].Value = newDate;
                                    worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    worksheet.Cells[i + lastRow, 2].Value = _monthlyCompliance[i].Zone;
                                    worksheet.Cells[i + lastRow, 3].Value = _monthlyCompliance[i].Observado;
                                    worksheet.Cells[i + lastRow, 4].Value = _monthlyCompliance[i].Compliance;
                                    worksheet.Cells[i + lastRow, 5].Value = _monthlyCompliance[i].Pit;
                                }
                                byte[] fileText2 = package.GetAsByteArray();
                                File.WriteAllBytes(loadFilePath, fileText2);
                                MyLastDateRefreshMonthlyValues = $"{StringResources.Updated}: {DateTime.Now}";
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

        private void GenerateDepressurizationTargetTemplate()
        {
            var months = new List<string> { "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            var places = new List<string> { "Pared Noreste Fuera Rajo", "Pared Noreste", "Pared Noreste Talud Bajo", "Pared Los Colorados", "Pared Los Colorados Talud Bajo", "Pared Este Fuera Rajo", "Pared Este Talud Medio" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = DepressurizationComplianceConstants.TargetDepressurizationWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var worksheet = excelPackage.Workbook.Worksheets.Add(DepressurizationComplianceConstants.TargetDepressurizationWorksheet);
            worksheet.Protection.IsProtected = true;

            for (var i = 0; i < 14; i++)            
                worksheet.Column(1 + i).Style.Locked = false;           

            worksheet.Cells["A2:A3"].Merge = true;
            worksheet.Cells["A2"].Value = "Depressurization Wall";
            worksheet.Cells["A2:B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A2:B2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B2"].Style.Font.Bold = true;
            worksheet.Column(1).Style.Font.Bold = true;
            worksheet.Column(1).Width = 27;

            for (var i = 0; i < places.Count; i++)            
                worksheet.Cells[4 + i, 1].Value = places[i];            

            worksheet.Cells["A4:A13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A4:A13"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(DepressurizationComplianceTemplateColors.HeaderBackgroundMonthlyDepressurizationCompliance));

            worksheet.Cells["B2:M2"].Merge = true;
            worksheet.Cells["B2"].Value = $"Targets (kPa) FY{MyFiscalYear}";
            worksheet.Cells["A2:A3"].Style.Font.Bold = true;
            worksheet.Cells["A2:M2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A2:M2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(DepressurizationComplianceTemplateColors.HeaderBackgroundMonthlyDepressurizationCompliance));

            worksheet.Row(3).Style.Font.Bold = true;
            for (var i = 0; i < months.Count; i++)
            {
                worksheet.Cells[3, 2 + i].Value = months[i];
                worksheet.Cells[3, 2 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[3, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(DepressurizationComplianceTemplateColors.HeaderBackgroundMonthlyDepressurizationCompliance));
                worksheet.Cells[3, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(2 + i).Width = 10;
            }

            for (var i = 0; i < 12; i++)
            {
                worksheet.Cells[$"A{1 + i}:M{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                for (var j = 0; j < 13; j++)                
                    worksheet.Cells[2 + i, 1 + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            worksheet.Cells[$"A13:M13"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = DepressurizationComplianceConstants.TargetDepressurizationExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledGenerateTargetTemplate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadDepressurizationTargetTemplate()
        {
            _targetCompliance.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var templateWorksheet = excelPackage.Workbook.Worksheets[DepressurizationComplianceConstants.TargetDepressurizationWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - DepressurizationComplianceConstants.TargetDepressurizationExcelFileName.Length) == DepressurizationComplianceConstants.TargetDepressurizationExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var openWriteCheck = File.OpenWrite(openFileDialog.FileName);
                        openWriteCheck.Close();

                        var _date = DateTime.Now;

                        for (var i = 0; i < 12; i++)
                        {
                            var _month = DateTime.ParseExact(templateWorksheet.Cells[3, 2 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;
                            _date = TemplateDates.ConvertDateToFiscalYearDate(i, MyFiscalYear, _month);

                            var rows = templateWorksheet.Dimension.Rows;

                            for (var j = 0; j < rows; j++)
                            {
                                if (templateWorksheet.Cells[4 + j, 1].Value != null)
                                {
                                    if (templateWorksheet.Cells[4 + j, 2 + i].Value == null)
                                    {
                                        templateWorksheet.Cells[4 + j, 2 + i].Value = -99000;
                                    }

                                    _targetCompliance.Add(new DepressurizationComplianceTargetCompliance()
                                    {
                                        Date = _date,
                                        Zone = templateWorksheet.Cells[4 + j, 1].Value.ToString(),
                                        Target = double.Parse(templateWorksheet.Cells[4 + j, 2 + i].Value.ToString())/1000
                                    });

                                }
                            }
                        }
                        excelPackage.Dispose();

                        var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.DepressurizationComplianceExcelFilePath;
                        var loadFileInfo = new FileInfo(loadFilePath);

                        if (loadFileInfo.Exists)
                        {
                            var package = new ExcelPackage(loadFileInfo);
                            var worksheet = package.Workbook.Worksheets[DepressurizationComplianceConstants.TargetDepressurizationSpotfireWorksheet];
                            if (worksheet != null)
                            {
                                var lastRow = worksheet.Dimension.End.Row + 1;
                                for (var i = 0; i < _targetCompliance.Count; i++)
                                {
                                    worksheet.Cells[i + lastRow, 1].Value = _targetCompliance[i].Date;
                                    worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    worksheet.Cells[i + lastRow, 2].Value = _targetCompliance[i].Zone;
                                    worksheet.Cells[i + lastRow, 3].Value = _targetCompliance[i].Target;
                                }
                                byte[] fileText2 = package.GetAsByteArray();
                                File.WriteAllBytes(loadFilePath, fileText2);
                                MyLastDateRefreshTargetValues = $"{StringResources.Updated}: {DateTime.Now}";
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
