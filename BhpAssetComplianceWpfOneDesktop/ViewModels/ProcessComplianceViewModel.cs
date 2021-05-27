using BhpAssetComplianceWpfOneDesktop.Resources;
using System;
using System.Collections.Generic;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Drawing;
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Models.ProcessComplianceModels;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class ProcessComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.ProcessCompliance;
        protected override string MyPosterIcon { get; set; } = IconKeys.ProcessCompliance;

        private string _myLastDateRefreshMonthlyValues;
        public string MyLastDateRefreshMonthlyValues
        {
            get { return _myLastDateRefreshMonthlyValues; }
            set { SetProperty(ref _myLastDateRefreshMonthlyValues, value); }
        }

        private DateTime _myMonthlyDate;
        public DateTime MyMonthlyDate
        {
            get { return _myMonthlyDate; }
            set { SetProperty(ref _myMonthlyDate, value); }
        }

        private bool _isEnabledLoadMonthlyValues;
        public bool IsEnabledLoadMonthlyValues
        {
            get { return _isEnabledLoadMonthlyValues; }
            set { SetProperty(ref _isEnabledLoadMonthlyValues, value); }
        }

        public DelegateCommand GenerateProcessComplianceTemplateCommand { get; private set; }
        public DelegateCommand LoadProcessComplianceTemplateCommand { get; private set; }

        private readonly List<ProcessComplianceOreToMill> _oreToMill = new List<ProcessComplianceOreToMill>();
        private readonly List<ProcessComplianceRecovery> _recovery = new List<ProcessComplianceRecovery>();
        private readonly List<ProcessComplianceOLAP> _OLAP = new List<ProcessComplianceOLAP>();
        private readonly List<ProcessComplianceSulphide> _sulphide = new List<ProcessComplianceSulphide>();

        public ProcessComplianceViewModel()
        {
            MyMonthlyDate = DateTime.Now;
            IsEnabledLoadMonthlyValues = false;
            GenerateProcessComplianceTemplateCommand = new DelegateCommand(GenerateProcessComplianceTemplate);
            LoadProcessComplianceTemplateCommand = new DelegateCommand(LoadProcessComplianceTemplate).ObservesCanExecute(() => IsEnabledLoadMonthlyValues);
        }

        private void GenerateProcessComplianceTemplate()
        {
            var phases1 = new List<string> { "Phase", "Ore to Mill Budget (Mt)", "Ore to Mill Actual (Mt)", "Hardness Budget (min)", "Hardness Actual (min)" };
            var phases2 = new List<string> { "Phase", "Recovery Budget (%)", "Recovery Actual (%)", "Feed Cu Budget (%)", "Feed Cu Actual (%)" };
            var feeds = new List<string> { "Feed Grade", "Stacked Ore (kt)", "CuT (%)", "Cathodes (t)", "Distribution", "Expit", "Average CuT", "Stocks" };
            var distributions = new List<string> { "Budget", "Actual", "Compliance %", "Budget %", "Actual %" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = ProcessComplianceConstants.ProcessComplianceWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var OreToMillWorksheet = excelPackage.Workbook.Worksheets.Add(ProcessComplianceConstants.OreToMillProcessComplianceWorksheet);
            OreToMillWorksheet.Cells["B1:C1"].Style.Font.Bold = true;
            OreToMillWorksheet.Cells["B1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OreToMillWorksheet.Cells["B1:C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
            OreToMillWorksheet.Cells["B1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            OreToMillWorksheet.Cells["B1"].Value = "Budget (min)";
            OreToMillWorksheet.Cells["C1"].Value = "Actual (min)";

            for (var i = 0; i < 3; i++)
            {
                OreToMillWorksheet.Cells[1, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                OreToMillWorksheet.Cells[2, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                OreToMillWorksheet.Cells[1, 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            OreToMillWorksheet.Cells["A2"].Style.Font.Bold = true;
            OreToMillWorksheet.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OreToMillWorksheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
            OreToMillWorksheet.Cells["A2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            OreToMillWorksheet.Cells["A2"].Value = "SPI Global";

            OreToMillWorksheet.Cells["A5:E5"].Style.Font.Bold = true;
            OreToMillWorksheet.Cells["A5:E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OreToMillWorksheet.Cells["A5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
            OreToMillWorksheet.Cells["B5:E5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
            OreToMillWorksheet.Cells["B5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < 17; i++)            
                OreToMillWorksheet.Cells[$"A{4 + i}:E{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;   

            string[] columns1 = { "A", "B", "C", "D", "E" };
            for (var i = columns1.GetLowerBound(0); i <= columns1.GetUpperBound(0); i++)
            {
                OreToMillWorksheet.Cells[5, 1 + i].Value = phases1[i];
                OreToMillWorksheet.Cells[$"{columns1[i]}5:{columns1[i]}20"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                OreToMillWorksheet.Column(1 + i).Width = 19;
            }
            OreToMillWorksheet.Column(1).Width = 11;

            var RecoveryWorksheet = excelPackage.Workbook.Worksheets.Add(ProcessComplianceConstants.RecoveryProcessComplianceWorksheet);
            RecoveryWorksheet.Cells["B1:C1"].Style.Font.Bold = true;
            RecoveryWorksheet.Cells["B1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            RecoveryWorksheet.Cells["B1:C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
            RecoveryWorksheet.Cells["B1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            RecoveryWorksheet.Cells["B1"].Value = "Budget (%)";
            RecoveryWorksheet.Cells["C1"].Value = "Actual (%)";

            for (var i = 0; i < 3; i++)
            {
                RecoveryWorksheet.Cells[1, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                RecoveryWorksheet.Cells[2, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                RecoveryWorksheet.Cells[1, 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            RecoveryWorksheet.Cells["A2"].Style.Font.Bold = true;
            RecoveryWorksheet.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            RecoveryWorksheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
            RecoveryWorksheet.Cells["A2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            RecoveryWorksheet.Cells["A2"].Value = "Rec Global";

            RecoveryWorksheet.Cells["A5:E5"].Style.Font.Bold = true;
            RecoveryWorksheet.Cells["A5:E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            RecoveryWorksheet.Cells["A5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
            RecoveryWorksheet.Cells["B5:E5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
            RecoveryWorksheet.Cells["B5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < 17; i++)
            {
                RecoveryWorksheet.Cells[$"A{4 + i}:E{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] columns2 = { "A", "B", "C", "D", "E" };
            for (var i = columns2.GetLowerBound(0); i <= columns2.GetUpperBound(0); i++)
            {
                RecoveryWorksheet.Cells[5, 1 + i].Value = phases2[i];
                RecoveryWorksheet.Cells[$"{columns2[i]}5:{columns2[i]}20"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                RecoveryWorksheet.Column(1 + i).Width = 19;
            }
            RecoveryWorksheet.Column(1).Width = 11;

            var OLAPWorksheet = excelPackage.Workbook.Worksheets.Add(ProcessComplianceConstants.OLAPProcessComplianceWorksheet);
            OLAPWorksheet.Cells["A1:H1"].Style.Font.Bold = true;
            OLAPWorksheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OLAPWorksheet.Cells["F1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OLAPWorksheet.Cells["A2:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OLAPWorksheet.Cells["F2:F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            OLAPWorksheet.Column(1).Width = 14;
            OLAPWorksheet.Column(6).Width = 14;

            string[] columns3 = { "A", "B", "C", "D", "F", "G", "H" };
            for (var i = 0; i < 4; i++)
            {
                OLAPWorksheet.Cells[1 + i, 1].Value = feeds[i];
                OLAPWorksheet.Cells[1 + i, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
                OLAPWorksheet.Cells[1 + i, 6].Value = feeds[4 + i];
                OLAPWorksheet.Cells[1 + i, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
                OLAPWorksheet.Column(2 + i).Width = 12;
                OLAPWorksheet.Column(7 + i).Width = 12;
                OLAPWorksheet.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                OLAPWorksheet.Cells[$"F{1 + i}:H{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                OLAPWorksheet.Cells[$"{columns3[i]}1:{columns3[i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            OLAPWorksheet.Cells["G1:H1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));

            for (var i = 0; i < 3; i++)
            {
                OLAPWorksheet.Cells[1, 2 + i].Value = distributions[i];
                OLAPWorksheet.Cells[1, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
                OLAPWorksheet.Cells[1, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                OLAPWorksheet.Cells[$"{columns3[4 + i]}1:{columns3[4 + i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            OLAPWorksheet.Cells[1, 7].Value = distributions[3];
            OLAPWorksheet.Cells[1, 8].Value = distributions[4];
            OLAPWorksheet.Cells["G1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var SulphideWorksheet = excelPackage.Workbook.Worksheets.Add(ProcessComplianceConstants.SulphideProcessComplianceWorksheet);
            SulphideWorksheet.Cells["A1:H1"].Style.Font.Bold = true;
            SulphideWorksheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SulphideWorksheet.Cells["F1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SulphideWorksheet.Cells["A2:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SulphideWorksheet.Cells["F2:F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SulphideWorksheet.Column(1).Width = 14;
            SulphideWorksheet.Column(6).Width = 14;

            string[] columns4 = { "A", "B", "C", "D", "F", "G", "H" };
            for (var i = 0; i < 4; i++)
            {
                SulphideWorksheet.Cells[1 + i, 1].Value = feeds[i];
                SulphideWorksheet.Cells[1 + i, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
                SulphideWorksheet.Cells[1 + i, 6].Value = feeds[4 + i];
                SulphideWorksheet.Cells[1 + i, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.OrangeBackgroundProcessCompliance));
                SulphideWorksheet.Column(2 + i).Width = 12;
                SulphideWorksheet.Column(7 + i).Width = 12;
                SulphideWorksheet.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                SulphideWorksheet.Cells[$"F{1 + i}:H{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                SulphideWorksheet.Cells[$"{columns4[i]}1:{columns4[i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            SulphideWorksheet.Cells["G1:H1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));

            for (var i = 0; i < 3; i++)
            {
                SulphideWorksheet.Cells[1, 2 + i].Value = distributions[i];
                SulphideWorksheet.Cells[1, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(ProcessComplianceTemplateColors.GrayBackgroundProcessCompliance));
                SulphideWorksheet.Cells[1, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                SulphideWorksheet.Cells[$"{columns4[4 + i]}1:{columns4[4 + i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            SulphideWorksheet.Cells[1, 7].Value = distributions[3];
            SulphideWorksheet.Cells[1, 8].Value = distributions[4];
            SulphideWorksheet.Cells["G1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = ProcessComplianceConstants.ProcessComplianceExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledLoadMonthlyValues = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }
        
        private void LoadProcessComplianceTemplate()
        {
            _oreToMill.Clear();
            _recovery.Clear();
            _OLAP.Clear();
            _sulphide.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var oreToMillTemplateWorksheet = excelPackage.Workbook.Worksheets[ProcessComplianceConstants.OreToMillProcessComplianceWorksheet];
                var recoveryTemplateWorksheet = excelPackage.Workbook.Worksheets[ProcessComplianceConstants.RecoveryProcessComplianceWorksheet];
                var OLAPTemplateWorksheet = excelPackage.Workbook.Worksheets[ProcessComplianceConstants.OLAPProcessComplianceWorksheet];
                var sulphideTemplateWorksheet = excelPackage.Workbook.Worksheets[ProcessComplianceConstants.SulphideProcessComplianceWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - ProcessComplianceConstants.ProcessComplianceExcelFileName.Length) == ProcessComplianceConstants.ProcessComplianceExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var fileStream = File.OpenWrite(openFileDialog.FileName);
                        fileStream.Close();

                        var rows = oreToMillTemplateWorksheet.Dimension.Rows;
                        for (var i = 0; i < rows; i++)
                        {
                            if (oreToMillTemplateWorksheet.Cells[i + 6, 1].Value != null)
                            {
                                if (oreToMillTemplateWorksheet.Cells[2, 2].Value == null)
                                    oreToMillTemplateWorksheet.Cells[2, 2].Value = -99;
                                if (oreToMillTemplateWorksheet.Cells[2, 3].Value == null)
                                    oreToMillTemplateWorksheet.Cells[2, 3].Value = -99;

                                for (var j = 0; j < 4; j++)
                                    if (oreToMillTemplateWorksheet.Cells[6 + i, 2 + j].Value == null)
                                        oreToMillTemplateWorksheet.Cells[6 + i, 2 + j].Value = -99;

                                _oreToMill.Add(new ProcessComplianceOreToMill()
                                {
                                    SpiGlobalBudget = double.Parse(oreToMillTemplateWorksheet.Cells[2, 2].Value.ToString()),
                                    SpiGlobalActual = double.Parse(oreToMillTemplateWorksheet.Cells[2, 3].Value.ToString()),
                                    Phase = oreToMillTemplateWorksheet.Cells[6 + i, 1].Value.ToString(),
                                    OretoMillBudget = double.Parse(oreToMillTemplateWorksheet.Cells[6 + i, 2].Value.ToString()),
                                    OretoMillActual = double.Parse(oreToMillTemplateWorksheet.Cells[6 + i, 3].Value.ToString()),
                                    HardnessBudget = double.Parse(oreToMillTemplateWorksheet.Cells[6 + i, 4].Value.ToString()),
                                    HardnessActual = double.Parse(oreToMillTemplateWorksheet.Cells[6 + i, 5].Value.ToString())
                                });
                            }
                        }

                        var rows2 = recoveryTemplateWorksheet.Dimension.Rows;
                        for (var i = 0; i < rows2; i++)
                        {
                            if (recoveryTemplateWorksheet.Cells[i + 6, 1].Value != null)
                            {
                                if (recoveryTemplateWorksheet.Cells[2, 2].Value == null)
                                    recoveryTemplateWorksheet.Cells[2, 2].Value = -99;

                                if (recoveryTemplateWorksheet.Cells[2, 3].Value == null)
                                    recoveryTemplateWorksheet.Cells[2, 3].Value = -99;

                                for (var j = 0; j < 4; j++)
                                    if (recoveryTemplateWorksheet.Cells[6 + i, 2 + j].Value == null)
                                        recoveryTemplateWorksheet.Cells[6 + i, 2 + j].Value = -99;

                                _recovery.Add(new ProcessComplianceRecovery()
                                {
                                    RecGlobalBudget = double.Parse(recoveryTemplateWorksheet.Cells[2, 2].Value.ToString()),
                                    RecGlobalActual = double.Parse(recoveryTemplateWorksheet.Cells[2, 3].Value.ToString()),
                                    Phase = recoveryTemplateWorksheet.Cells[6 + i, 1].Value.ToString(),
                                    RecoveryBudget = double.Parse(recoveryTemplateWorksheet.Cells[6 + i, 2].Value.ToString())/100,
                                    RecoveryActual = double.Parse(recoveryTemplateWorksheet.Cells[6 + i, 3].Value.ToString())/100,
                                    FeedCuBudget = double.Parse(recoveryTemplateWorksheet.Cells[6 + i, 4].Value.ToString())/100,
                                    FeedCuActual = double.Parse(recoveryTemplateWorksheet.Cells[6 + i, 5].Value.ToString())/100
                                });
                            }
                        }

                        for (var i = 0; i < 3; i++)
                        {
                            for (var j = 0; j < 3; j++)
                                if (OLAPTemplateWorksheet.Cells[2 + i, 2 + j].Value == null)
                                    OLAPTemplateWorksheet.Cells[2 + i, 2 + j].Value = -99;

                            for (var j = 0; j < 2; j++)
                                if (OLAPTemplateWorksheet.Cells[2 + i, 7 + j].Value == null)
                                    OLAPTemplateWorksheet.Cells[2 + i, 7 + j].Value = -99;

                            _OLAP.Add(new ProcessComplianceOLAP()
                            {
                                FeedGrade = OLAPTemplateWorksheet.Cells[2 + i, 1].Value.ToString(),
                                Budget = double.Parse(OLAPTemplateWorksheet.Cells[2 + i, 2].Value.ToString()),
                                Actual = double.Parse(OLAPTemplateWorksheet.Cells[2 + i, 3].Value.ToString()),
                                Compliance = double.Parse(OLAPTemplateWorksheet.Cells[2 + i, 4].Value.ToString())/100,
                                Distribution = OLAPTemplateWorksheet.Cells[2 + i, 6].Value.ToString(),
                                DistributionBudget = double.Parse(OLAPTemplateWorksheet.Cells[2 + i, 7].Value.ToString())/100,
                                DistributionActual = double.Parse(OLAPTemplateWorksheet.Cells[2 + i, 8].Value.ToString())/100
                            });
                        }

                        for (var i = 0; i < 3; i++)
                        {
                            for (var j = 0; j < 3; j++)
                                if (sulphideTemplateWorksheet.Cells[2 + i, 2 + j].Value == null)
                                    sulphideTemplateWorksheet.Cells[2 + i, 2 + j].Value = -99;

                            for (var j = 0; j < 2; j++)
                                if (sulphideTemplateWorksheet.Cells[2 + i, 7 + j].Value == null)
                                    sulphideTemplateWorksheet.Cells[2 + i, 7 + j].Value = -99;

                            _sulphide.Add(new ProcessComplianceSulphide()
                            {
                                FeedGrade = sulphideTemplateWorksheet.Cells[2 + i, 1].Value.ToString(),
                                Budget = double.Parse(sulphideTemplateWorksheet.Cells[2 + i, 2].Value.ToString()),
                                Actual = double.Parse(sulphideTemplateWorksheet.Cells[2 + i, 3].Value.ToString()),
                                Compliance = double.Parse(sulphideTemplateWorksheet.Cells[2 + i, 4].Value.ToString())/100,
                                Distribution = sulphideTemplateWorksheet.Cells[2 + i, 6].Value.ToString(),
                                DistributionBudget = double.Parse(sulphideTemplateWorksheet.Cells[2 + i, 7].Value.ToString())/100,
                                DistributionActual = double.Parse(sulphideTemplateWorksheet.Cells[2 + i, 8].Value.ToString())/100
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



                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.ProcessComplianceExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);


                if (loadFileInfo.Exists)
                {
                   try
                   {
                        var package = new ExcelPackage(loadFileInfo);
                        var oreToMillWorksheet = package.Workbook.Worksheets[ProcessComplianceConstants.OreToMillProcessComplianceWorksheet];
                        var recoveryWorksheet = package.Workbook.Worksheets[ProcessComplianceConstants.RecoveryProcessComplianceWorksheet];
                        var OLAPWorksheet = package.Workbook.Worksheets[ProcessComplianceConstants.OLAPProcessComplianceWorksheet];
                        var sulphideWorksheet = package.Workbook.Worksheets[ProcessComplianceConstants.SulphideProcessComplianceWorksheet];

                        if (oreToMillWorksheet != null & recoveryWorksheet != null & OLAPWorksheet != null & sulphideWorksheet != null)
                        {
                            var newDate = new DateTime(MyMonthlyDate.Year, MyMonthlyDate.Month, 1, 00, 00, 00);
                            var lastRow1 = oreToMillWorksheet.Dimension.End.Row + 1;
                            for (var i = 0; i < _oreToMill.Count; i++)
                            {
                                oreToMillWorksheet.Cells[i + lastRow1, 1].Value = newDate;
                                oreToMillWorksheet.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                oreToMillWorksheet.Cells[i + lastRow1, 2].Value = _oreToMill[i].SpiGlobalBudget;
                                oreToMillWorksheet.Cells[i + lastRow1, 3].Value = _oreToMill[i].SpiGlobalActual;
                                oreToMillWorksheet.Cells[i + lastRow1, 4].Value = _oreToMill[i].Phase;
                                oreToMillWorksheet.Cells[i + lastRow1, 5].Value = _oreToMill[i].OretoMillBudget;
                                oreToMillWorksheet.Cells[i + lastRow1, 6].Value = _oreToMill[i].OretoMillActual;
                                oreToMillWorksheet.Cells[i + lastRow1, 7].Value = _oreToMill[i].HardnessBudget;
                                oreToMillWorksheet.Cells[i + lastRow1, 8].Value = _oreToMill[i].HardnessActual;
                            }

                            var lastRow2 = recoveryWorksheet.Dimension.End.Row + 1;
                            for (var i = 0; i < _recovery.Count; i++)
                            {
                                recoveryWorksheet.Cells[i + lastRow2, 1].Value = newDate;
                                recoveryWorksheet.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                recoveryWorksheet.Cells[i + lastRow2, 2].Value = _recovery[i].RecGlobalBudget;
                                recoveryWorksheet.Cells[i + lastRow2, 3].Value = _recovery[i].RecGlobalActual;
                                recoveryWorksheet.Cells[i + lastRow2, 4].Value = _recovery[i].Phase;
                                recoveryWorksheet.Cells[i + lastRow2, 5].Value = _recovery[i].RecoveryBudget;
                                recoveryWorksheet.Cells[i + lastRow2, 6].Value = _recovery[i].RecoveryActual;
                                recoveryWorksheet.Cells[i + lastRow2, 7].Value = _recovery[i].FeedCuBudget;
                                recoveryWorksheet.Cells[i + lastRow2, 8].Value = _recovery[i].FeedCuActual;
                            }

                            var lastRow3 = OLAPWorksheet.Dimension.End.Row + 1;
                            for (var i = 0; i < _OLAP.Count; i++)
                            {
                                OLAPWorksheet.Cells[i + lastRow3, 1].Value = newDate;
                                OLAPWorksheet.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                OLAPWorksheet.Cells[i + lastRow3, 2].Value = _OLAP[i].FeedGrade;
                                OLAPWorksheet.Cells[i + lastRow3, 3].Value = _OLAP[i].Budget;
                                OLAPWorksheet.Cells[i + lastRow3, 4].Value = _OLAP[i].Actual;
                                OLAPWorksheet.Cells[i + lastRow3, 5].Value = _OLAP[i].Compliance;
                                OLAPWorksheet.Cells[i + lastRow3, 6].Value = _OLAP[i].Distribution;
                                OLAPWorksheet.Cells[i + lastRow3, 7].Value = _OLAP[i].DistributionBudget;
                                OLAPWorksheet.Cells[i + lastRow3, 8].Value = _OLAP[i].DistributionActual;
                            }

                            var lastRow4 = sulphideWorksheet.Dimension.End.Row + 1;
                            for (var i = 0; i < _sulphide.Count; i++)
                            {
                                sulphideWorksheet.Cells[i + lastRow4, 1].Value = newDate;
                                sulphideWorksheet.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                sulphideWorksheet.Cells[i + lastRow4, 2].Value = _sulphide[i].FeedGrade;
                                sulphideWorksheet.Cells[i + lastRow4, 3].Value = _sulphide[i].Budget;
                                sulphideWorksheet.Cells[i + lastRow4, 4].Value = _sulphide[i].Actual;
                                sulphideWorksheet.Cells[i + lastRow4, 5].Value = _sulphide[i].Compliance;
                                sulphideWorksheet.Cells[i + lastRow4, 6].Value = _sulphide[i].Distribution;
                                sulphideWorksheet.Cells[i + lastRow4, 7].Value = _sulphide[i].DistributionBudget;
                                sulphideWorksheet.Cells[i + lastRow4, 8].Value = _sulphide[i].DistributionActual;
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
                   catch (Exception ex)
                   {
                        MessageBox.Show(ex.Message, StringResources.UploadError);
                   }

                    //if (oreToMillWorksheet != null & recoveryWorksheet != null & OLAPWorksheet != null & sulphideWorksheet != null)
                    //{
                    //    try
                    //    {
                    //        //var openWriteCheck = File.OpenWrite(loadFilePath);
                    //        //openWriteCheck.Close();

                            
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        MessageBox.Show(ex.Message, StringResources.UploadError);
                    //    }
                    //}
                    //else
                    //{
                    //    var wrongFileMessage = $"{StringResources.WorksheetNotExist} {loadFilePath} {StringResources.IsTheRightOne}";
                    //    MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                    //}                   
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

