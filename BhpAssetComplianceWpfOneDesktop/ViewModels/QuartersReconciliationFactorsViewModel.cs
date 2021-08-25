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
using BhpAssetComplianceWpfOneDesktop.Engines;
using BhpAssetComplianceWpfOneDesktop.Utility;
using BhpAssetComplianceWpfOneDesktop.Models.QuartersReconciliationFactorsModels;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class QuartersReconciliationFactorsViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.QuartersReconciliationFactors;
        protected override string MyPosterIcon { get; set; } = IconKeys.ReconciliationFactors;

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

        public DelegateCommand GenerateQuartersReconciliationFactorsTemplateCommand { get; private set; }
        public DelegateCommand LoadQuartersReconciliationFactorsTemplateCommand { get; private set; }

        public QuartersReconciliationFactorsViewModel()
        {
            MyMonthlyDate = DateTime.Now;
            IsEnabledLoadMonthlyValues = false;
            GenerateQuartersReconciliationFactorsTemplateCommand =
                new DelegateCommand(GenerateQuartersReconciliationFactorsTemplate);
            LoadQuartersReconciliationFactorsTemplateCommand =
                new DelegateCommand(LoadQuartersReconciliationFactorsTemplate).ObservesCanExecute(() =>
                    IsEnabledLoadMonthlyValues);
        }

        private void GenerateQuartersReconciliationFactorsTemplate()
        {
            var headers = new List<string> { "F0", "F1", "F2", "F3" };
            var elements = new List<string>
                { "Ore", "%CuT", "Cu Fines", "Ore", "%CuT", "Cu Fines", "Ore", "%CuT", "Cu Fines", "Cu Fines" };
            var quarters = new List<string>
                { "Mill", "Q1", "Q2", "Q3", "Q4", "OL", "Q1", "Q2", "Q3", "Q4", "SL", "Q1", "Q2", "Q3", "Q4" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title =
                QuartersReconciliationFactorsConstants.QuartersReconciliationFactorsWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var worksheet =
                excelPackage.Workbook.Worksheets.Add(QuartersReconciliationFactorsConstants
                    .QuartersReconciliationFactorsWorksheet);
            worksheet.Column(2).Width = 11;
            int[] rows1 = { 7, 18, 29 };
            int[] rows2 = { 10, 21, 32 };

            var _date = TemplateDates.ConvertDateToFiscalYearString(MyMonthlyDate);

            for (var i = rows1.GetLowerBound(0); i <= rows1.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"B{rows1[i]}:B{rows2[i]}"].Merge = true;
                worksheet.Cells[rows1[i], 2].Style.Font.Bold = true;
                worksheet.Cells[rows1[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rows1[i], 2].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .DarkGrayBackgroundMineCompliance));
                worksheet.Cells[rows1[i], 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[rows1[i], 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"B{rows1[i]}"].Value = _date;
                worksheet.Cells[$"B{rows1[i]}:B{rows1[i] + 3}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                worksheet.Cells[$"B{rows1[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            }

            worksheet.Column(3).Width = 13;
            worksheet.Column(3).Style.Font.Bold = true;
            worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B10:M10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["B21:L21"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["B32:I32"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["C3:M3"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["C14:L14"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["C25:I25"].Style.Border.Top.Style = ExcelBorderStyle.Thick;

            int[] rows3 = { 3, 14, 25 };
            int[] rows4 = { 6, 17, 28 };
            for (var i = rows3.GetLowerBound(0); i <= rows3.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"C{rows3[i]}:C{rows4[i]}"].Merge = true;
                worksheet.Cells[rows3[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rows3[i], 3].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .DarkGrayBackgroundMineCompliance));
                worksheet.Cells[$"D{7 + i}:M{7 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{18 + i}:L{18 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{29 + i}:I{29 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            worksheet.Cells["C3"].Value = quarters[0];
            worksheet.Cells["C14"].Value = quarters[5];
            worksheet.Cells["C25"].Value = quarters[10];
            string[] columns1 = { "D", "E", "F", "G", "H" };
            for (var i = 0; i < 5; i++)
            {
                worksheet.Cells[6 + i, 3].Value = quarters[i];
                worksheet.Cells[17 + i, 3].Value = quarters[i + 5];
                worksheet.Cells[28 + i, 3].Value = quarters[i + 10];
                worksheet.Cells[$"{columns1[i]}25:{columns1[i]}32"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            string[] columns2 = { "D", "E", "F", "G", "H", "I", "J", "K" };
            for (var i = columns2.GetLowerBound(0); i <= columns2.GetUpperBound(0); i++)
                worksheet.Cells[$"{columns2[i]}14:{columns2[i]}21"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            string[] columns3 = { "B", "C" };

            for (var i = columns3.GetLowerBound(0); i <= columns3.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"{columns3[i]}3:{columns3[i]}10"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                worksheet.Cells[$"{columns3[i]}14:{columns3[i]}21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                worksheet.Cells[$"{columns3[i]}25:{columns3[i]}32"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                worksheet.Cells[$"D{25 + i}:F{25 + i}"].Merge = true;
                worksheet.Cells[$"G{25 + i}:I{25 + i}"].Merge = true;

                worksheet.Cells[5 + i, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[5 + i, 13].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .LightOrangeBackgroundMineCompliance));
                worksheet.Cells[$"J{5 + i}:L{5 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"J{5 + i}:L{5 + i}"].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .LightGreenBackgroundMineCompliance));
                worksheet.Cells[$"J{16 + i}:L{16 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"J{16 + i}:L{16 + i}"].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .LightGreenBackgroundMineCompliance));
            }

            worksheet.Cells["M3:M10"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["L14:L21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["I25:I32"].Style.Border.Right.Style = ExcelBorderStyle.Thick;


            worksheet.Cells["C7:C10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["C7:C10"].Style.Fill.BackgroundColor.SetColor(
                ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.GrayBackgroundMineCompliance));
            worksheet.Cells["C18:C21"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["C18:C21"].Style.Fill.BackgroundColor.SetColor(
                ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.GrayBackgroundMineCompliance));
            worksheet.Cells["C29:C32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["C29:C32"].Style.Fill.BackgroundColor.SetColor(
                ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.GrayBackgroundMineCompliance));

            worksheet.Cells["C6:C9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["C17:C20"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["C28:C31"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            int[] rows5 = { 3, 4, 14, 15 };
            for (var i = rows5.GetLowerBound(0); i <= rows5.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"D{rows5[i]}:F{rows5[i]}"].Merge = true;
                worksheet.Cells[$"G{rows5[i]}:I{rows5[i]}"].Merge = true;
                worksheet.Cells[$"J{rows5[i]}:L{rows5[i]}"].Merge = true;
                worksheet.Cells[$"D{3 + i}:M{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{14 + i}:L{14 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{25 + i}:I{25 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            int[] rows6 = { 5, 6, 16, 17, 27, 28 };
            for (var i = rows6.GetLowerBound(0); i <= rows6.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"D{rows6[i]}:F{rows6[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"D{rows6[i]}:F{rows6[i]}"].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .LightGreenBackgroundMineCompliance));
                worksheet.Cells[$"G{rows6[i]}:I{rows6[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"G{rows6[i]}:I{rows6[i]}"].Style.Fill.BackgroundColor.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                        .LightOrangeBackgroundMineCompliance));

                worksheet.Cells[27, 4 + i].Value = elements[i];
                worksheet.Cells[27, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[28, 4 + i].Value = "%";
                worksheet.Cells[28, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            string[] cells = { "D3", "G3", "J3", "M3", "D14", "G14", "J14", "G25", "D25" };
            string[] cells2 = { "D4", "G4", "J4", "M4", "D15", "G15", "J15", "G26", "D26" };
            string[] columns4 = { "D", "E", "F", "G", "H", "I", "J", "K", "L" };

            for (var i = columns4.GetLowerBound(0); i <= columns4.GetUpperBound(0); i++)
                worksheet.Cells[$"{columns4[i]}3:{columns4[i]}10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (var i = cells.GetLowerBound(0); i <= cells.GetUpperBound(0); i++)
            {
                worksheet.Cells[16, 4 + i].Value = elements[i];
                worksheet.Cells[17, 4 + i].Value = "%";
                worksheet.Cells[16, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[17, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i % 2 != 0)
                {
                    worksheet.Cells[cells[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cells[i]].Style.Fill.BackgroundColor.SetColor(
                        ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                            .DarkOrangeBackgroundMineCompliance));
                    worksheet.Cells[cells2[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cells2[i]].Style.Fill.BackgroundColor.SetColor(
                        ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                            .OrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    worksheet.Cells[cells[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cells[i]].Style.Fill.BackgroundColor.SetColor(
                        ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                            .DarkGreenBackgroundMineCompliance));
                    worksheet.Cells[cells2[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cells2[i]].Style.Fill.BackgroundColor.SetColor(
                        ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors
                            .GreenBackgroundMineCompliance));
                }
            }

            string[] cells3 = { "D3", "G3", "J3", "M3" };
            string[] cells4 = { "D4", "G4", "J4", "M4" };

            for (var i = cells3.GetLowerBound(0); i <= cells3.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"{cells3[i]}"].Value = headers[i];
                worksheet.Cells[$"{cells3[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells3[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"{cells4[i]}"].Value = "Quarter";
                worksheet.Cells[$"{cells4[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells4[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (var i = 0; i < 10; i++)
            {
                worksheet.Cells[5, 4 + i].Value = elements[i];
                worksheet.Cells[5, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[6, 4 + i].Value = "%";
                worksheet.Cells[6, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(4 + i).Width = 11;
            }

            string[] cells5 = { "D14", "G14", "J14" };
            string[] cells6 = { "D15", "G15", "J15" };
            for (var i = cells5.GetLowerBound(0); i <= cells5.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"{cells5[i]}"].Value = headers[i];
                worksheet.Cells[$"{cells5[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells5[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"{cells6[i]}"].Value = "Month/Year";
                worksheet.Cells[$"{cells6[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells6[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            string[] cells7 = { "D25", "G25" };
            string[] cells8 = { "D26", "G26" };
            for (var i = cells7.GetLowerBound(0); i <= cells7.GetUpperBound(0); i++)
            {
                worksheet.Cells[$"{cells7[i]}"].Value = headers[i];
                worksheet.Cells[$"{cells7[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells7[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"{cells8[i]}"].Value = "Month/Year";
                worksheet.Cells[$"{cells8[i]}"].Style.Font.Color.SetColor(
                    ColorTranslator.FromHtml(QuartersReconciliationFactorsTemplateColors.WhiteFontMineCompliance));
                worksheet.Cells[$"{cells8[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = QuartersReconciliationFactorsConstants.QuartersReconciliationFactorExcelFileName,
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

        private void LoadQuartersReconciliationFactorsTemplate()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            var result = openFileDialog.ShowDialog();
            if (!result.HasValue || !result.Value) return;

            var quartersReconciliationFactorsEngine =
                new QuartersReconciliationFactorsReadTemplate(openFileDialog.FileName);

            QuarterReconciliationFactors quarterReconciliationFactors;
            try
            {
                quarterReconciliationFactors = quartersReconciliationFactorsEngine.Process();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, StringResources.UploadError);
                return;
            }

            // if (openFileDialog.ShowDialog() == true)
            // {
            //     var openFilePath = new FileInfo(openFileDialog.FileName);
            //     var excelPackage = new ExcelPackage(openFilePath);
            //     var worksheet =
            //         excelPackage.Workbook.Worksheets[
            //             QuartersReconciliationFactorsConstants.QuartersReconciliationFactorsWorksheet];
            //
            //     if (openFilePath.FullName.Substring(openFilePath.FullName.Length -
            //                                         QuartersReconciliationFactorsConstants
            //                                             .QuartersReconciliationFactorExcelFileName.Length) ==
            //         QuartersReconciliationFactorsConstants.QuartersReconciliationFactorExcelFileName)
            //     {
            //         try
            //         {
            //             // Check if the file is already open
            //             var fileStream = File.OpenWrite(openFileDialog.FileName);
            //             fileStream.Close();
            //
            //             for (var i = 0; i < 4; i++)
            //             {
            //                 for (var j = 0; j < 10; j++)
            //                     if (worksheet.Cells[7 + i, 4 + j].Value == null)
            //                         worksheet.Cells[7 + i, 4 + j].Value = -99;
            //
            //                 for (var j = 0; j < 9; j++)
            //                     if (worksheet.Cells[18 + i, 4 + j].Value == null)
            //                         worksheet.Cells[18 + i, 4 + j].Value = -99;
            //
            //                 for (var j = 0; j < 6; j++)
            //                     if (worksheet.Cells[29 + i, 4 + j].Value == null)
            //                         worksheet.Cells[29 + i, 4 + j].Value = -99;
            //             }
            //
            //             for (var i = 0; i < 4; i++)
            //             {
            //                 _F0.Add(new QuartersReconciliationFactorsF0()
            //                 {
            //                     Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
            //                     MillOre = double.Parse(worksheet.Cells[7 + i, 4].Value.ToString()) / 100,
            //                     OLOre = double.Parse(worksheet.Cells[18 + i, 4].Value.ToString()) / 100,
            //                     SLOre = double.Parse(worksheet.Cells[29 + i, 4].Value.ToString()) / 100,
            //                     MillCuT = double.Parse(worksheet.Cells[7 + i, 5].Value.ToString()) / 100,
            //                     OLCuT = double.Parse(worksheet.Cells[18 + i, 5].Value.ToString()) / 100,
            //                     SLCuT = double.Parse(worksheet.Cells[29 + i, 5].Value.ToString()) / 100,
            //                     MillCuFines = double.Parse(worksheet.Cells[7 + i, 6].Value.ToString()) / 100,
            //                     OLCuFines = double.Parse(worksheet.Cells[18 + i, 6].Value.ToString()) / 100,
            //                     SLCuFines = double.Parse(worksheet.Cells[29 + i, 6].Value.ToString()) / 100
            //                 });
            //             }
            //
            //             for (var i = 0; i < 4; i++)
            //             {
            //                 _F1.Add(new QuartersReconciliationFactorsF1()
            //                 {
            //                     Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
            //                     MillOre = double.Parse(worksheet.Cells[7 + i, 7].Value.ToString()) / 100,
            //                     OLOre = double.Parse(worksheet.Cells[18 + i, 7].Value.ToString()) / 100,
            //                     SLOre = double.Parse(worksheet.Cells[29 + i, 7].Value.ToString()) / 100,
            //                     MillCuT = double.Parse(worksheet.Cells[7 + i, 8].Value.ToString()) / 100,
            //                     OLCuT = double.Parse(worksheet.Cells[18 + i, 8].Value.ToString()) / 100,
            //                     SLCuT = double.Parse(worksheet.Cells[29 + i, 8].Value.ToString()) / 100,
            //                     MillCuFines = double.Parse(worksheet.Cells[7 + i, 9].Value.ToString()) / 100,
            //                     OLCuFines = double.Parse(worksheet.Cells[18 + i, 9].Value.ToString()) / 100,
            //                     SLCuFines = double.Parse(worksheet.Cells[29 + i, 9].Value.ToString()) / 100
            //                 });
            //             }
            //
            //             for (var i = 0; i < 4; i++)
            //             {
            //                 _F2.Add(new QuartersReconciliationFactorsF2()
            //                 {
            //                     Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
            //                     MillOre = double.Parse(worksheet.Cells[7 + i, 10].Value.ToString()) / 100,
            //                     OLOre = double.Parse(worksheet.Cells[18 + i, 10].Value.ToString()) / 100,
            //                     MillCuT = double.Parse(worksheet.Cells[7 + i, 11].Value.ToString()) / 100,
            //                     OLCuT = double.Parse(worksheet.Cells[18 + i, 11].Value.ToString()) / 100,
            //                     MillCuFines = double.Parse(worksheet.Cells[7 + i, 12].Value.ToString()) / 100,
            //                     OLCuFines = double.Parse(worksheet.Cells[18 + i, 12].Value.ToString()) / 100
            //                 });
            //             }
            //
            //             for (var i = 0; i < 4; i++)
            //             {
            //                 _F3.Add(new QuartersReconciliationFactorsF3()
            //                 {
            //                     Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
            //                     MillCuFines = double.Parse(worksheet.Cells[7 + i, 13].Value.ToString()) / 100
            //                 });
            //             }
            //
            //             excelPackage.Dispose();
            //         }
            //         catch (Exception ex)
            //         {
            //             MessageBox.Show(ex.Message, StringResources.UploadError);
            //         }
            //     }
            //     else
            //     {
            //         var wrongFileMessage =
            //             $"{StringResources.WrongUploadedFile} {openFilePath.FullName} {StringResources.IsTheRightOne}";
            //         MessageBox.Show(wrongFileMessage, StringResources.UploadError);
            //     }
            //
            //     var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default
            //         .QuartersReconciliationFactorsExcelFilePath;
            //     var loadFileInfo = new FileInfo(loadFilePath);
            //
            //     if (loadFileInfo.Exists)
            //     {
            //         var package = new ExcelPackage(loadFileInfo);
            //         var F0Worksheet = package.Workbook.Worksheets[
            //             QuartersReconciliationFactorsConstants.F0QuartersReconciliationFactorsSpotfireWorksheet];
            //         var F1Worsheet = package.Workbook.Worksheets[
            //             QuartersReconciliationFactorsConstants.F1QuartersReconciliationFactorsSpotfireWorksheet];
            //         var F2Worksheet = package.Workbook.Worksheets[
            //             QuartersReconciliationFactorsConstants.F2QuartersReconciliationFactorsSpotfireWorksheet];
            //         var F3Woksheet = package.Workbook.Worksheets[
            //             QuartersReconciliationFactorsConstants.F3QuartersReconciliationFactorsSpotfireWorksheet];
            //
            //         if (F0Worksheet != null & F1Worsheet != null & F2Worksheet != null & F3Woksheet != null)
            //         {
            //             try
            //             {
            //                 var openWriteCheck = File.OpenWrite(loadFilePath);
            //                 openWriteCheck.Close();
            //
            //                 var newDate = new DateTime(MyMonthlyDate.Year, MyMonthlyDate.Month, 1, 00, 00, 00);
            //                 var lastRow1 = F0Worksheet.Dimension.End.Row + 1;
            //                 for (var i = 0; i < _F0.Count; i++)
            //                 {
            //                     F0Worksheet.Cells[i + lastRow1, 1].Value = newDate;
            //                     F0Worksheet.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            //                     F0Worksheet.Cells[i + lastRow1, 2].Value = _F0[i].Quarter;
            //                     F0Worksheet.Cells[i + lastRow1, 3].Value = _F0[i].MillOre;
            //                     F0Worksheet.Cells[i + lastRow1, 4].Value = _F0[i].OLOre;
            //                     F0Worksheet.Cells[i + lastRow1, 5].Value = _F0[i].SLOre;
            //                     F0Worksheet.Cells[i + lastRow1, 6].Value = _F0[i].MillCuT;
            //                     F0Worksheet.Cells[i + lastRow1, 7].Value = _F0[i].OLCuT;
            //                     F0Worksheet.Cells[i + lastRow1, 8].Value = _F0[i].SLCuT;
            //                     F0Worksheet.Cells[i + lastRow1, 9].Value = _F0[i].MillCuFines;
            //                     F0Worksheet.Cells[i + lastRow1, 10].Value = _F0[i].OLCuFines;
            //                     F0Worksheet.Cells[i + lastRow1, 11].Value = _F0[i].SLCuFines;
            //                 }
            //
            //                 var lastRow2 = F1Worsheet.Dimension.End.Row + 1;
            //                 for (var i = 0; i < _F1.Count; i++)
            //                 {
            //                     F1Worsheet.Cells[i + lastRow2, 1].Value = newDate;
            //                     F1Worsheet.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            //                     F1Worsheet.Cells[i + lastRow2, 2].Value = _F1[i].Quarter;
            //                     F1Worsheet.Cells[i + lastRow2, 3].Value = _F1[i].MillOre;
            //                     F1Worsheet.Cells[i + lastRow2, 4].Value = _F1[i].OLOre;
            //                     F1Worsheet.Cells[i + lastRow2, 5].Value = _F1[i].SLOre;
            //                     F1Worsheet.Cells[i + lastRow2, 6].Value = _F1[i].MillCuT;
            //                     F1Worsheet.Cells[i + lastRow2, 7].Value = _F1[i].OLCuT;
            //                     F1Worsheet.Cells[i + lastRow2, 8].Value = _F1[i].SLCuT;
            //                     F1Worsheet.Cells[i + lastRow2, 9].Value = _F1[i].MillCuFines;
            //                     F1Worsheet.Cells[i + lastRow2, 10].Value = _F1[i].OLCuFines;
            //                     F1Worsheet.Cells[i + lastRow2, 11].Value = _F1[i].SLCuFines;
            //                 }
            //
            //                 var lastRow3 = F2Worksheet.Dimension.End.Row + 1;
            //                 for (var i = 0; i < _F2.Count; i++)
            //                 {
            //                     F2Worksheet.Cells[i + lastRow3, 1].Value = newDate;
            //                     F2Worksheet.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            //                     F2Worksheet.Cells[i + lastRow3, 2].Value = _F2[i].Quarter;
            //                     F2Worksheet.Cells[i + lastRow3, 3].Value = _F2[i].MillOre;
            //                     F2Worksheet.Cells[i + lastRow3, 4].Value = _F2[i].OLOre;
            //                     F2Worksheet.Cells[i + lastRow3, 5].Value = _F2[i].MillCuT;
            //                     F2Worksheet.Cells[i + lastRow3, 6].Value = _F2[i].OLCuT;
            //                     F2Worksheet.Cells[i + lastRow3, 7].Value = _F2[i].MillCuFines;
            //                     F2Worksheet.Cells[i + lastRow3, 8].Value = _F2[i].OLCuFines;
            //                 }
            //
            //                 var lastRow4 = F3Woksheet.Dimension.End.Row + 1;
            //                 for (var i = 0; i < _F3.Count; i++)
            //                 {
            //                     F3Woksheet.Cells[i + lastRow4, 1].Value = newDate;
            //                     F3Woksheet.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            //                     F3Woksheet.Cells[i + lastRow4, 2].Value = _F3[i].Quarter;
            //                     F3Woksheet.Cells[i + lastRow4, 3].Value = _F3[i].MillCuFines;
            //                 }
            //
            //                 byte[] fileText2 = package.GetAsByteArray();
            //                 File.WriteAllBytes(loadFilePath, fileText2);
            //                 MyLastDateRefreshMonthlyValues = $"{StringResources.Updated}: {DateTime.Now}";
            //             }
            //             catch (Exception ex)
            //             {
            //                 MessageBox.Show(ex.Message, StringResources.UploadError);
            //             }
            //         }
            //         else
            //         {
            //             var wrongFileMessage =
            //                 $"{StringResources.WorksheetNotExist} {loadFilePath} {StringResources.IsTheRightOne}";
            //             MessageBox.Show(wrongFileMessage, StringResources.UploadError);
            //         }
            //     }
            //     else
            //     {
            //         var wrongFileMessage =
            //             $"{StringResources.WorksheetNotExist} {loadFilePath} {StringResources.ExistsOrNotSelect}";
            //         MessageBox.Show(wrongFileMessage, StringResources.UploadError);
            //     }
        }
    }
}