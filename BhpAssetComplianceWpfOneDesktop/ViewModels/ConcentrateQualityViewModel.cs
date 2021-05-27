using BhpAssetComplianceWpfOneDesktop.Constants;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Models.ConcentrateQualityModels;
using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    //TODO: BORRAME
    public class ConcentrateQualityViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.ConcentrateQuality;
        protected override string MyPosterIcon { get; set; } = IconKeys.ConcentrateQuality;

        private string _myLastDateRefreshActualValues;
        public string MyLastDateRefreshActualValues
        {
            get { return _myLastDateRefreshActualValues; }
            set { SetProperty(ref _myLastDateRefreshActualValues, value); }
        }

        private string _myLastDateRefreshBudgetValues;
        public string MyLastDateRefreshBudgetValues
        {
            get { return _myLastDateRefreshBudgetValues; }
            set { SetProperty(ref _myLastDateRefreshBudgetValues, value); }
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

        private bool _isEnabledLoadActualValues;
        public bool IsEnabledLoadActualValues
        {
            get { return _isEnabledLoadActualValues; }
            set { SetProperty(ref _isEnabledLoadActualValues, value); }
        }

        private bool _isEnabledLoadBudgetValues;
        public bool IsEnabledLoadBudgetValues
        {
            get { return _isEnabledLoadBudgetValues; }
            set { SetProperty(ref _isEnabledLoadBudgetValues, value); }
        }

        public DelegateCommand GenerateActualFreightTemplateCommand { get; set; }
        public DelegateCommand LoadActualFreightTemplateCommand { get; set; }
        public DelegateCommand GenerateBudgetFreightTemplateCommand { get; set; }                            
        public DelegateCommand LoadBudgetFreightTemplateCommand { get; set; }


        private readonly List<ConcentrateQualityActualFreight> _actualFreights = new List<ConcentrateQualityActualFreight>();
        private readonly List<ConcentrateQualityBudgetFreight> _budgetFreights = new List<ConcentrateQualityBudgetFreight>();

        public ConcentrateQualityViewModel()
        {
            MyDateActual = DateTime.Now;
            MyFiscalYear = MyDateActual.Year;
            IsEnabledLoadActualValues = false;
            IsEnabledLoadBudgetValues = false;
            GenerateActualFreightTemplateCommand = new DelegateCommand(GenerateActualFreightTemplate);
            LoadActualFreightTemplateCommand = new DelegateCommand(LoadActualFreightTemplate).ObservesCanExecute(() => IsEnabledLoadActualValues);
            GenerateBudgetFreightTemplateCommand = new DelegateCommand(GenerateBudgetFreightTemplate);
            LoadBudgetFreightTemplateCommand = new DelegateCommand(LoadBudgetFreightTemplate).ObservesCanExecute(() => IsEnabledLoadBudgetValues);
        }

        private void GenerateActualFreightTemplate()
        {
            var headers = new List<string> { "Nombre M/N", "N°", "Inicio embarque", "Termino embarque" };
            var items = new List<string> { "WMT", "DMT", "Moisture.", "Cu", "As", "Fe", "Au", "Ag", "S", "Insol.", "Cd", "Zn", "Hg", "SiO2", "Al2O3", "Sb", "Mo" };
            var units = new List<string> { "Pesometer t", "Pesometer t", "%", "%", "%", "%", "g/t", "g/t", "%", "%", "%", "%", "g/t", "%", "%", "%", "%" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = ConcentrateQualityConstants.ActualFreightWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var worksheet = excelPackage.Workbook.Worksheets.Add(ConcentrateQualityConstants.ActualFreightWorksheet);
            worksheet.Cells["A2:U2"].Style.Font.Bold = true;
            worksheet.Cells["A2:U2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.FontActualConcentrateQuality));  // TODO: Utilizar constantes para colores
            worksheet.Cells["E3:U3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.FontActualConcentrateQuality));
            worksheet.Cells["A2:U2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["E3:U3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(1).Width = 22;

            worksheet.Cells["A2:A3"].Merge = true;
            worksheet.Cells["B2:B3"].Merge = true;
            worksheet.Cells["C2:C3"].Merge = true;
            worksheet.Cells["D2:D3"].Merge = true;

            for (var i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[2, i + 1].Value = headers[i];
                worksheet.Cells[2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[2, i + 1].Style.WrapText = true;
            }

            for (var i = 0; i < items.Count; i++)
            {
                worksheet.Cells[2, i + 5].Value = items[i];
                worksheet.Cells[3, i + 5].Value = units[i];
            }

            for (var i = 1; i < 22; i++)
            {
                if (i % 2 != 0)
                {
                    worksheet.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.HeaderBackgroundActualConcentrateQuality1));
                }
                else if (i % 2 == 0)
                {
                    worksheet.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.HeaderBackgroundActualConcentrateQuality2));
                }
                worksheet.Column(1 + i).Width = 16;
            }

            for (var i = 1; i < 18; i++)
            {
                if (i % 2 != 0)
                {
                    worksheet.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.UnitsBackgroundActualConcentrateQuality1));
                }
                else if (i % 2 == 0)
                {
                    worksheet.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.UnitsBackgroundActualConcentrateQuality2));
                }
            }

            worksheet.Cells["A2:U2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < 14; i++)
            {
                worksheet.Cells[i + 2, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = ConcentrateQualityConstants.ActualFreightExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledLoadActualValues = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadActualFreightTemplate()
        {
            _actualFreights.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var templateWorksheet = excelPackage.Workbook.Worksheets[ConcentrateQualityConstants.ActualFreightWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - ConcentrateQualityConstants.ActualFreightExcelFileName.Length) == ConcentrateQualityConstants.ActualFreightExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var fileStream = File.OpenWrite(openFileDialog.FileName);
                        fileStream.Close();

                        var rows = templateWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows; i++)
                        {
                            if (templateWorksheet.Cells[i + 3, 1].Value != null)
                            {
                                for (var j = 0; j < 17; j++)
                                    if (templateWorksheet.Cells[3 + i, 5 + j].Value == null)
                                        templateWorksheet.Cells[3 + i, 5 + j].Value = -99;

                                _actualFreights.Add(new ConcentrateQualityActualFreight()
                                {
                                    Name = templateWorksheet.Cells[3 + i, 1].Value.ToString(),
                                    Number = Int32.Parse(templateWorksheet.Cells[3 + i, 2].Value.ToString()),
                                    Start = Convert.ToDateTime(templateWorksheet.Cells[3 + i, 3].Value.ToString()),
                                    End = Convert.ToDateTime(templateWorksheet.Cells[3 + i, 4].Value.ToString()),
                                    WMT = double.Parse(templateWorksheet.Cells[3 + i, 5].Value.ToString()),
                                    DMT = double.Parse(templateWorksheet.Cells[3 + i, 6].Value.ToString()),
                                    Moisture = double.Parse(templateWorksheet.Cells[3 + i, 7].Value.ToString())/100,
                                    Cu = double.Parse(templateWorksheet.Cells[3 + i, 8].Value.ToString())/100,
                                    As = double.Parse(templateWorksheet.Cells[3 + i, 9].Value.ToString()) * 10000,
                                    Fe = double.Parse(templateWorksheet.Cells[3 + i, 10].Value.ToString())/100,
                                    Au = double.Parse(templateWorksheet.Cells[3 + i, 11].Value.ToString()),
                                    Ag = double.Parse(templateWorksheet.Cells[3 + i, 12].Value.ToString()),
                                    S = double.Parse(templateWorksheet.Cells[3 + i, 13].Value.ToString())/100,
                                    Insoluble = double.Parse(templateWorksheet.Cells[3 + i, 14].Value.ToString())/100,
                                    Cd = double.Parse(templateWorksheet.Cells[3 + i, 15].Value.ToString()) * 10000,
                                    Zn = double.Parse(templateWorksheet.Cells[3 + i, 16].Value.ToString()) * 10000,
                                    Hg = double.Parse(templateWorksheet.Cells[3 + i, 17].Value.ToString()),
                                    SiO2 = double.Parse(templateWorksheet.Cells[3 + i, 18].Value.ToString())/100,
                                    Al2O3 = double.Parse(templateWorksheet.Cells[3 + i, 19].Value.ToString())/100,
                                    Sb = double.Parse(templateWorksheet.Cells[3 + i, 20].Value.ToString()) * 10000,
                                    Mo = double.Parse(templateWorksheet.Cells[3 + i, 21].Value.ToString())/100 
                                });
                            }
                        }
                        excelPackage.Dispose();
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
                
                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ConcentrateQualityExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);

                if (loadFileInfo.Exists)
                {
                    var package = new ExcelPackage(loadFileInfo);
                    var worksheet = package.Workbook.Worksheets[ConcentrateQualityConstants.ActualFreightSpotfireWorksheet];
                    if (worksheet != null)
                    {
                        try
                        {
                            var openWriteCheck = File.OpenWrite(loadFilePath);
                            openWriteCheck.Close();

                            var lastRow = worksheet.Dimension.End.Row + 1;
                            var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);

                            for (var i = 0; i < _actualFreights.Count; i++)
                            {
                                worksheet.Cells[i + lastRow, 1].Value = newDate;
                                worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                worksheet.Cells[i + lastRow, 2].Value = _actualFreights[i].Name;
                                worksheet.Cells[i + lastRow, 3].Value = _actualFreights[i].Number;
                                worksheet.Cells[i + lastRow, 4].Value = _actualFreights[i].Start;
                                worksheet.Cells[i + lastRow, 4].Style.Numberformat.Format = "yyyy-MM-dd";
                                worksheet.Cells[i + lastRow, 5].Value = _actualFreights[i].End;
                                worksheet.Cells[i + lastRow, 5].Style.Numberformat.Format = "yyyy-MM-dd";
                                worksheet.Cells[i + lastRow, 6].Value = _actualFreights[i].WMT;
                                worksheet.Cells[i + lastRow, 7].Value = _actualFreights[i].DMT;
                                worksheet.Cells[i + lastRow, 8].Value = _actualFreights[i].Moisture;
                                worksheet.Cells[i + lastRow, 9].Value = _actualFreights[i].Cu;
                                worksheet.Cells[i + lastRow, 10].Value = _actualFreights[i].As;
                                worksheet.Cells[i + lastRow, 11].Value = _actualFreights[i].Fe;
                                worksheet.Cells[i + lastRow, 12].Value = _actualFreights[i].Au;
                                worksheet.Cells[i + lastRow, 13].Value = _actualFreights[i].Ag;
                                worksheet.Cells[i + lastRow, 14].Value = _actualFreights[i].S;
                                worksheet.Cells[i + lastRow, 15].Value = _actualFreights[i].Insoluble;
                                worksheet.Cells[i + lastRow, 16].Value = _actualFreights[i].Cd;
                                worksheet.Cells[i + lastRow, 17].Value = _actualFreights[i].Zn;
                                worksheet.Cells[i + lastRow, 18].Value = _actualFreights[i].Hg;
                                worksheet.Cells[i + lastRow, 19].Value = _actualFreights[i].SiO2;
                                worksheet.Cells[i + lastRow, 20].Value = _actualFreights[i].Al2O3;
                                worksheet.Cells[i + lastRow, 21].Value = _actualFreights[i].Sb;
                                worksheet.Cells[i + lastRow, 22].Value = _actualFreights[i].Mo;
                            }
                            byte[] fileText2 = package.GetAsByteArray();
                            File.WriteAllBytes(loadFilePath, fileText2);
                            MyLastDateRefreshActualValues = $"{StringResources.Updated}: {DateTime.Now}";
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
        }

        private void GenerateBudgetFreightTemplate()
        {
            var headers = new List<string> { "Item", "Unit", "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            var items = new List<string> { "Au", "Ag", "Mo", "As", "Cd", "Pb", "Zn", "Bi", "Sb", "Fe Conc", "Fe", "Py Conc", "Py", "S2", "Concentrate Grade" };
            var units = new List<string> { "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "%", "%", "%", "%", "%", "%" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = ConcentrateQualityConstants.BudgetFreightWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var worksheet = excelPackage.Workbook.Worksheets.Add(ConcentrateQualityConstants.BudgetFreightWorksheet);

            worksheet.Cells["B2:O2"].Merge = true;
            worksheet.Cells["B2"].Value = $"FY{MyFiscalYear}";

            worksheet.Column(2).Style.Font.Bold = true;
            worksheet.Column(3).Style.Font.Bold = true;

            worksheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.HeaderBackgroundBudgetConcentrateQuality));
            worksheet.Cells["B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[3, i + 2].Value = headers[i];
                worksheet.Cells[3, i + 2].Style.Font.Bold = true;
                worksheet.Cells[3, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[3, i + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.ItemsBackgroundBudgetConcentrateQuality));
            }

            worksheet.Cells["B2:O2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Column(2).Width = 22;
            worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < 17; i++)
            {
                worksheet.Cells[i + 2, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"B{2 + i}:O{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"B{2 + i}:O{2 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Column(3 + i).Width = 11;
                worksheet.Cells[3, i + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (var i = 0; i < items.Count; i++)
            {
                worksheet.Cells[4 + i, 2].Value = items[i];
                worksheet.Cells[4 + i, 3].Value = units[i];
                worksheet.Cells[$"B{4 + i}:C{4 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"B{4 + i}:C{4 + i}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ConcentrateQualityTemplateColors.ItemsBackgroundBudgetConcentrateQuality));
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = ConcentrateQualityConstants.BudgetFreightExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledLoadBudgetValues = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadBudgetFreightTemplate()
        {
            _budgetFreights.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var templateWorksheet = excelPackage.Workbook.Worksheets[ConcentrateQualityConstants.BudgetFreightWorksheet]; // TODO: Llevar a un archivo de constantes

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - ConcentrateQualityConstants.BudgetFreightExcelFileName.Length) == ConcentrateQualityConstants.BudgetFreightExcelFileName)
                {
                    try
                    {
                        var fileStream = File.OpenWrite(openFileDialog.FileName);
                        fileStream.Close();

                        var _date = DateTime.Now;

                        for (var i = 0; i < 12; i++)
                        {
                            var _month = DateTime.ParseExact(templateWorksheet.Cells[3, 4 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;
                            _date = TemplateDates.ConvertDateToFiscalYearDate(i, MyFiscalYear, _month);

                            for (var j = 0; j < 17; j++)
                            {
                                if (templateWorksheet.Cells[4 + j, 4 + i].Value == null)
                                    templateWorksheet.Cells[4 + j, 4 + i].Value = -99;
                            }

                            _budgetFreights.Add(new ConcentrateQualityBudgetFreight()
                            {
                                Date = _date,
                                Au = double.Parse(templateWorksheet.Cells[4, 4 + i].Value.ToString()),
                                Ag = double.Parse(templateWorksheet.Cells[5, 4 + i].Value.ToString()),
                                Mo = double.Parse(templateWorksheet.Cells[6, 4 + i].Value.ToString()) / 1000000,
                                As = double.Parse(templateWorksheet.Cells[7, 4 + i].Value.ToString()),
                                Cd = double.Parse(templateWorksheet.Cells[8, 4 + i].Value.ToString()),
                                Pb = double.Parse(templateWorksheet.Cells[9, 4 + i].Value.ToString()),
                                Zn = double.Parse(templateWorksheet.Cells[10, 4 + i].Value.ToString()),
                                Bi = double.Parse(templateWorksheet.Cells[11, 4 + i].Value.ToString()),
                                Sb = double.Parse(templateWorksheet.Cells[12, 4 + i].Value.ToString()),
                                FeConc = double.Parse(templateWorksheet.Cells[13, 4 + i].Value.ToString())/100,
                                Fe = double.Parse(templateWorksheet.Cells[14, 4 + i].Value.ToString())/100,
                                PyConc = double.Parse(templateWorksheet.Cells[15, 4 + i].Value.ToString())/100,
                                Py = double.Parse(templateWorksheet.Cells[16, 4 + i].Value.ToString())/100,
                                S2 = double.Parse(templateWorksheet.Cells[17, 4 + i].Value.ToString())/100,
                                ConcentrateGrade = double.Parse(templateWorksheet.Cells[18, 4 + i].Value.ToString())/100
                            });
                        }
                        excelPackage.Dispose();
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

                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ConcentrateQualityExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);
                if (loadFileInfo.Exists)
                {
                    var package = new ExcelPackage(loadFileInfo);
                    var worksheet = package.Workbook.Worksheets[ConcentrateQualityConstants.BudgetFreightWorksheet];
                    if (worksheet != null)
                    {
                        try
                        {
                            var openWriteCheck = File.OpenWrite(loadFilePath);
                            openWriteCheck.Close();

                            var lastRow = worksheet.Dimension.End.Row + 1;

                            for (var i = 0; i < _budgetFreights.Count; i++)
                            {
                                worksheet.Cells[i + lastRow, 1].Value = _budgetFreights[i].Date;
                                worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                worksheet.Cells[i + lastRow, 2].Value = _budgetFreights[i].Au;
                                worksheet.Cells[i + lastRow, 3].Value = _budgetFreights[i].Ag;
                                worksheet.Cells[i + lastRow, 4].Value = _budgetFreights[i].Mo;
                                worksheet.Cells[i + lastRow, 5].Value = _budgetFreights[i].As;
                                worksheet.Cells[i + lastRow, 6].Value = _budgetFreights[i].Cd;
                                worksheet.Cells[i + lastRow, 7].Value = _budgetFreights[i].Pb;
                                worksheet.Cells[i + lastRow, 8].Value = _budgetFreights[i].Zn;
                                worksheet.Cells[i + lastRow, 9].Value = _budgetFreights[i].Bi;
                                worksheet.Cells[i + lastRow, 10].Value = _budgetFreights[i].Sb;
                                worksheet.Cells[i + lastRow, 11].Value = _budgetFreights[i].FeConc;
                                worksheet.Cells[i + lastRow, 12].Value = _budgetFreights[i].Fe;
                                worksheet.Cells[i + lastRow, 13].Value = _budgetFreights[i].PyConc;
                                worksheet.Cells[i + lastRow, 14].Value = _budgetFreights[i].Py;
                                worksheet.Cells[i + lastRow, 15].Value = _budgetFreights[i].S2;
                                worksheet.Cells[i + lastRow, 16].Value = _budgetFreights[i].ConcentrateGrade;
                            }

                            byte[] fileText2 = package.GetAsByteArray();
                            File.WriteAllBytes(loadFilePath, fileText2);
                            MyLastDateRefreshBudgetValues = $"{StringResources.Updated}: {DateTime.Now}";
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
        }
    }
}
