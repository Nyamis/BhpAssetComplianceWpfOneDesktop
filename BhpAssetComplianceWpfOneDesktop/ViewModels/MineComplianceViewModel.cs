using BhpAssetComplianceWpfOneDesktop.Resources;
using System;
using System.Collections.Generic;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Drawing;
using System.Globalization;
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Extensions;
using BhpAssetComplianceWpfOneDesktop.Models.MineComplianceModels;
using BhpAssetComplianceWpfOneDesktop.Utility;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class MineComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.MineCompliance;
        protected override string MyPosterIcon { get; set; } = IconKeys.MineMovement;

        private string _myLastDateRefreshRealValues;
        public string MyLastDateRefreshRealValues
        {
            get { return _myLastDateRefreshRealValues; }
            set { SetProperty(ref _myLastDateRefreshRealValues, value); }
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

        private bool _isEnabledGenerateRealTemplate;
        public bool IsEnabledGenerateRealTemplate
        {
            get { return _isEnabledGenerateRealTemplate; }
            set { SetProperty(ref _isEnabledGenerateRealTemplate, value); }
        }

        private bool _isEnabledGenerateBudgetTemplate;
        public bool IsEnabledGenerateBudgetTemplate
        {
            get { return _isEnabledGenerateBudgetTemplate; }
            set { SetProperty(ref _isEnabledGenerateBudgetTemplate, value); }
        }

        public DelegateCommand GenerateMineComplianceRealTemplateCommand { get; private set; }
        public DelegateCommand LoadMineComplianceRealTemplateCommand { get; private set; }
        public DelegateCommand GenerateMineComplianceBudgetTemplateCommand { get; private set; }
        public DelegateCommand LoadMineComplianceBudgetTemplateCommand { get; private set; }

        private readonly List<MineComplianceRealMovementProduction> _realMovementProduction = new List<MineComplianceRealMovementProduction>();
        private readonly List<MineComplianceRealPitDisintegrated> _realPitDisintegrated = new List<MineComplianceRealPitDisintegrated>();
        private readonly List<MineComplianceRealMillFc> _realMillFc = new List<MineComplianceRealMillFc>();
        private readonly List<MineComplianceRealLoadingFc> _realLoadingFc = new List<MineComplianceRealLoadingFc>();
        private readonly List<MineComplianceRealHaulingFc> _realHaulingFc = new List<MineComplianceRealHaulingFc>();
        private readonly List<MineComplianceBudgetPrincipal> _budgetPrincipal = new List<MineComplianceBudgetPrincipal>();
        private readonly List<MineComplianceBudgetMovementProduction> _budgetMovementProduction = new List<MineComplianceBudgetMovementProduction>();
        private readonly List<MineComplianceBudgetPitDisintegrated> _budgetPitDisintegrated = new List<MineComplianceBudgetPitDisintegrated>();

        public MineComplianceViewModel()
        {
            MyDateActual = DateTime.Now;
            MyFiscalYear = MyDateActual.Year;
            IsEnabledGenerateRealTemplate = false;
            IsEnabledGenerateBudgetTemplate = false;
            GenerateMineComplianceRealTemplateCommand = new DelegateCommand(GenerateMineComplianceRealTemplate);
            LoadMineComplianceRealTemplateCommand = new DelegateCommand(LoadMineComplianceRealTemplate).ObservesCanExecute(() => IsEnabledGenerateRealTemplate);
            GenerateMineComplianceBudgetTemplateCommand = new DelegateCommand(GenerateMineComplianceBudgetTemplate);
            LoadMineComplianceBudgetTemplateCommand = new DelegateCommand(LoadMineComplianceBudgetTemplate).ObservesCanExecute(() => IsEnabledGenerateBudgetTemplate);
        }

        private void GenerateMineComplianceRealTemplate()
        {
            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = MineComplianceConstants.RealMineComplianceWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var realPitDisintegratedWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.RealPitDisintegratedMineComplianceWorksheet);

            var categories = new List<string> { "Category", "Parameter", "Unit", "Month" };
            var pitDisintegratedParameters = new List<string> { "Expit ES", "Expit EN", "Mill", "OL", "SL", "Other", "Reh", "Mov." };

            realPitDisintegratedWorksheet.Cells["C2:J2"].Merge = true;
            realPitDisintegratedWorksheet.Cells["C2:J2"].Style.Font.Bold = true;
            realPitDisintegratedWorksheet.Cells["C2:J2"].Value = "Real";
            realPitDisintegratedWorksheet.Cells["C2:J2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            realPitDisintegratedWorksheet.Cells["C2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realPitDisintegratedWorksheet.Cells["C2:J2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["C2:J2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkBlueBackgroundMineCompliance));

            realPitDisintegratedWorksheet.Cells[$"C2:J2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            realPitDisintegratedWorksheet.Cells[$"B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            realPitDisintegratedWorksheet.Cells[$"J2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < 5; i++)
            {
                realPitDisintegratedWorksheet.Cells[$"B{2 + i}:J{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] columns1 = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
            for (var i = columns1.GetLowerBound(0); i <= columns1.GetUpperBound(0); i++)
            {
                realPitDisintegratedWorksheet.Cells[$"{columns1[i]}3:{columns1[i]}6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                realPitDisintegratedWorksheet.Column(1 + i).Width = 14;
            }

            for (var i = 0; i < categories.Count; i++)
            {
                realPitDisintegratedWorksheet.Cells[3 + i, 2].Value = categories[i];
                realPitDisintegratedWorksheet.Cells[3 + i, 2].Style.Font.Bold = true;
                realPitDisintegratedWorksheet.Cells[3 + i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                realPitDisintegratedWorksheet.Cells[3 + i, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            }

            realPitDisintegratedWorksheet.Cells["C3:D3"].Merge = true;
            realPitDisintegratedWorksheet.Cells["E3:H3"].Merge = true;
            realPitDisintegratedWorksheet.Cells["I3:J3"].Merge = true;
            realPitDisintegratedWorksheet.Cells["C3:J3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            realPitDisintegratedWorksheet.Cells[3, 3].Value = "Movement";
            realPitDisintegratedWorksheet.Cells[3, 5].Value = "Rehandling";
            realPitDisintegratedWorksheet.Cells[3, 9].Value = "Total";

            realPitDisintegratedWorksheet.Cells["C3:J3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["C3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));
            realPitDisintegratedWorksheet.Cells["E3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["E3:H3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

            for (var i = 0; i < 3; i++)
            {
                realPitDisintegratedWorksheet.Cells[$"C{3 + i}:J{3 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (var i = 0; i < pitDisintegratedParameters.Count; i++)
            {
                realPitDisintegratedWorksheet.Cells[4, 3 + i].Value = pitDisintegratedParameters[i];
                realPitDisintegratedWorksheet.Cells[5, 3 + i].Value = "t";
            }
            realPitDisintegratedWorksheet.Cells["C4:J4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            realPitDisintegratedWorksheet.Cells["C4:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["C4:J4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
            realPitDisintegratedWorksheet.Cells["E4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["E4:H4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

            realPitDisintegratedWorksheet.Cells["C5:J5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["C5:J5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
            realPitDisintegratedWorksheet.Cells["E5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realPitDisintegratedWorksheet.Cells["E5:H5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));

            var realMovementProductionWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.RealMovementProductionMineComplianceWorksheet);

            var zones = new List<string> { "Los Colorados", "Laguna Seca", "Laguna Seca 2", "Oxide", "Sulphide Leach", "Coloso" };
            var movementProductionParameters1 = new List<string> { "Ore Grade - CuT", "Mill Recovery", "Mill Feed", "Cu Ex Mill", "Runtime", "Hours" };
            var movementProductionParameters2 = new List<string> { "CuT Ore Grade (ROM)", "CuS Ore Grade (ROM)", "CuT Ore Grade (Crusher)", "CuS Ore Grade (Crusher)", "Recovery (Crusher + ROM)", "ROM", "Crushed Material", "Total Stacked Material", "Cu Cathodes", "CuT Ore Grade", "CuS Ore Grade", "Recovery", "Stacked Material from Mine", "Contractors Stacked Material from Stocks", "MEL Stacked Material from Stocks", "Total Stacked Material", "Cu Cathodes" };
            var movementProductionUnits1 = new List<string> { "%", "%", "t", "t Cu", "%", "h" };
            var movementProductionUnits2 = new List<string> { "%", "%", "%", "%", "%", "dmt", "dmt", "dmt", "t Cu", "%", "%", "%", "dmt", "dmt", "dmt", "dmt", "t Cu" };

            realMovementProductionWorksheet.Column(1).Width = 19;
            int[] rows1 = { 3, 11, 19, 27, 36, 46 };
            for (var i = rows1.GetLowerBound(0); i <= rows1.GetUpperBound(0); i++)
            {
                realMovementProductionWorksheet.Cells[rows1[i], 1].Value = zones[i];
                realMovementProductionWorksheet.Cells[$"B{rows1[i]}:D{rows1[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

                if (i % 2 != 0)
                {
                    realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));
                }
            }

            realMovementProductionWorksheet.Column(2).Width = 35;
            realMovementProductionWorksheet.Column(2).Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            for (var i = 0; i < movementProductionParameters1.Count; i++)
            {
                realMovementProductionWorksheet.Cells[3 + i, 2].Value = movementProductionParameters1[i];
                realMovementProductionWorksheet.Cells[3 + i, 3].Value = movementProductionUnits1[i];
                realMovementProductionWorksheet.Cells[$"B{3 + i}:D{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                realMovementProductionWorksheet.Cells[11 + i, 2].Value = movementProductionParameters1[i];
                realMovementProductionWorksheet.Cells[11 + i, 3].Value = movementProductionUnits1[i];
                realMovementProductionWorksheet.Cells[$"B{11 + i}:D{11 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                realMovementProductionWorksheet.Cells[19 + i, 2].Value = movementProductionParameters1[i];
                realMovementProductionWorksheet.Cells[19 + i, 3].Value = movementProductionUnits1[i];
                realMovementProductionWorksheet.Cells[$"B{19 + i}:D{19 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            int[] rows2 = { 3, 19, 36 };
            int[] rows3 = { 8, 24, 43 };
            for (var i = rows2.GetLowerBound(0); i <= rows2.GetUpperBound(0); i++)
            {
                realMovementProductionWorksheet.Cells[$"B{rows2[i]}:B{rows3[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                realMovementProductionWorksheet.Cells[$"B{rows2[i]}:B{rows3[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
                realMovementProductionWorksheet.Cells[$"B{rows2[i]}:B{rows3[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"B{rows2[i]}:B{rows3[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                realMovementProductionWorksheet.Cells[$"C{rows2[i]}:C{rows3[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                realMovementProductionWorksheet.Cells[$"C{rows2[i]}:C{rows3[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
                realMovementProductionWorksheet.Cells[$"C{rows2[i]}:C{rows3[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"C{rows2[i]}:C{rows3[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"D{rows2[i]}:D{rows3[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            int[] rows4 = { 11, 27, 46 };
            int[] rows5 = { 16, 35, 47 };
            for (var i = rows4.GetLowerBound(0); i <= rows4.GetUpperBound(0); i++)
            {
                realMovementProductionWorksheet.Cells[$"B{rows4[i]}:B{rows5[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                realMovementProductionWorksheet.Cells[$"B{rows4[i]}:B{rows5[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));
                realMovementProductionWorksheet.Cells[$"B{rows4[i]}:B{rows5[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"B{rows4[i]}:B{rows5[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                realMovementProductionWorksheet.Cells[$"C{rows4[i]}:C{rows5[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                realMovementProductionWorksheet.Cells[$"C{rows4[i]}:C{rows5[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
                realMovementProductionWorksheet.Cells[$"C{rows4[i]}:C{rows5[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"C{rows4[i]}:C{rows5[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[$"D{rows4[i]}:D{rows5[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            realMovementProductionWorksheet.Column(3).Width = 11;
            realMovementProductionWorksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realMovementProductionWorksheet.Column(4).Width = 11;

            int[] rows6 = { 2, 10, 18, 26, 45 };
            for (var i = 0; i < 5; i++)
            {
                realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
                realMovementProductionWorksheet.Cells[rows6[i], 3].Value = "Unit";
                realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                if (i % 2 != 0)
                {
                    realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realMovementProductionWorksheet.Cells[rows6[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
                }
            }
            realMovementProductionWorksheet.Cells[45, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMovementProductionWorksheet.Cells[45, 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

            for (var i = 0; i < movementProductionUnits2.Count; i++)
            {
                realMovementProductionWorksheet.Cells[27 + i, 3].Value = movementProductionUnits2[i];
                realMovementProductionWorksheet.Cells[27 + i, 2].Value = movementProductionParameters2[i];
                realMovementProductionWorksheet.Cells[$"B{27 + i}:D{27 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            realMovementProductionWorksheet.Cells[46, 2].Value = "Cu Ex Coloso";
            realMovementProductionWorksheet.Cells[46, 3].Value = "t";
            realMovementProductionWorksheet.Cells["B46:D46"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            realMovementProductionWorksheet.Cells[47, 2].Value = "Cu from low grade Concentrate";
            realMovementProductionWorksheet.Cells[47, 3].Value = "t";
            realMovementProductionWorksheet.Cells["B47:D47"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            realMovementProductionWorksheet.Cells[2, 4].Value = "Month";
            realMovementProductionWorksheet.Cells[2, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMovementProductionWorksheet.Cells[2, 4].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));

            realMovementProductionWorksheet.Cells["D2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            realMovementProductionWorksheet.Cells["D2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            realMovementProductionWorksheet.Cells["D2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            realMovementProductionWorksheet.Cells["D2"].Style.Font.Bold = true;
            realMovementProductionWorksheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var realLoadingWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.RealLoadingFCMineComplianceWorksheet);
           
            var realLoadingParameters = new List<string> { "Units", "Availability", "Utilization (New TUM)", "Total Hours", "Available Hours", "Equipment scheduled downtime", "Equipment non-scheduled downtime", "Process scheduled downtime", "Process non-scheduled downtime", "Stand By", "Hang Time", "Production time(New TUM)", "Performance (New TUM)", "Total Tonnes" };
            var realLoadingHeaders = new List<string> { "SHOVEL BUCYRUS 495B", "SHOVEL P & H 4100XPB", "SHOVEL P & H 4100XPC", "SHOVEL BUCYRUS 495HR", "FRONT LOADER CAT 994F", "PC5500", "PC8000", "SUMMARY SHOVEL 73 yd3", "TOTAL FLEET" };
            var realLoadingUnits = new List<string> { "N°", "%", "%", "h", "h", "h", "h", "h", "h", "h", "h", "h", "t/h", "kt" };

            realLoadingWorksheet.Column(1).Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            realLoadingWorksheet.Column(1).Width = 36;
            realLoadingWorksheet.Cells["A1:A2"].Merge = true;
            realLoadingWorksheet.Cells["A1"].Value = "Total Loading Fleet";
            realLoadingWorksheet.Cells["A1:C1"].Style.Font.Bold = true;
            realLoadingWorksheet.Cells["A1"].Style.Font.Size = 16;
            realLoadingWorksheet.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.BlackFontMineCompliance));
            realLoadingWorksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            realLoadingWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realLoadingWorksheet.Column(2).Width = 16;

            realLoadingWorksheet.Cells["C4"].Style.Font.Bold = true;
            realLoadingWorksheet.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realLoadingWorksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realLoadingWorksheet.Cells["C4"].Value = "Month";
 
            realLoadingWorksheet.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realLoadingWorksheet.Cells["C4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            realLoadingWorksheet.Cells["C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            realLoadingWorksheet.Cells["C4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            int[] rows7 = { 4, 20, 36, 52, 68, 84, 100, 116, 132 };
            for (var i = rows7.GetLowerBound(0); i <= rows7.GetUpperBound(0); i++)
            {
                realLoadingWorksheet.Cells[rows7[i], 1].Value = realLoadingHeaders[i];
                realLoadingWorksheet.Cells[rows7[i], 1].Style.Font.Bold = true;
                realLoadingWorksheet.Cells[rows7[i], 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                realLoadingWorksheet.Cells[rows7[i], 2].Value = "Unit";
                realLoadingWorksheet.Cells[rows7[i], 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

                realLoadingWorksheet.Cells[$"A{rows7[i]}:B{rows7[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                realLoadingWorksheet.Cells[$"A{rows7[i]}:c{rows7[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                realLoadingWorksheet.Cells[$"A{rows7[i]}:B{rows7[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                if (i % 2 != 0)
                {
                    realLoadingWorksheet.Cells[rows7[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[rows7[i] + 14, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i] + 14, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[rows7[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 }:A{rows7[i] + 13 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 }:A{rows7[i] + 13 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    realLoadingWorksheet.Cells[rows7[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[rows7[i] + 14, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i] + 14, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[rows7[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[rows7[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 }:A{rows7[i] + 13 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 }:A{rows7[i] + 13 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
                }

                for (var j = 0; j < realLoadingParameters.Count; j++)
                {
                    realLoadingWorksheet.Cells[rows7[i] + 1 + j, 1].Value = realLoadingParameters[j];
                    realLoadingWorksheet.Cells[rows7[i] + 1 + j, 2].Value = realLoadingUnits[j];
                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 + j}:C{rows7[i] + 1 + j}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    realLoadingWorksheet.Cells[$"A{rows7[i] + 1 + j}:C{rows7[i] + 1 + j}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    if (i % 2 != 0)
                    {
                        realLoadingWorksheet.Cells[rows7[i] + 1 + j, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        realLoadingWorksheet.Cells[rows7[i] + 1 + j, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
                    }
                    else if (i % 2 == 0)
                    {
                        realLoadingWorksheet.Cells[rows7[i] + 1 + j, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        realLoadingWorksheet.Cells[rows7[i] + 1 + j, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
                    }

                }
            }
            realLoadingWorksheet.Column(3).Width = 11;

            var realHaulingWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.RealHaulingFCMineComplianceWorksheet);
                      
            var realHaulingParameters = new List<string> { "Units", "Mechanical Availability", "Physical Availability", "Utilization (New TUM)", "Total Hours", "Available Hours", "Equipment scheduled downtime", "Equipment non-scheduled downtime", "Process scheduled downtime", "Process non-scheduled downtime", "Standby", "Queue Time", "Production time (New TUM)", "Performance (New TUM)", "Cycle Time", "Total Tonnes" };
            var realHaulingHeaders = new List<string> { "930 Fleet", "960 Autonomous Fleet", "960 MEL Fleet", "Liebherr ESTRS", "Komatsu ESTRS", "CAT ESTRS", "797B MARC Fleet", "797B MEL Fleet", "797F MARC Fleet", "793F MEL Fleet", "240 Total Fleet (793 + 930)", "350 Total Fleet (797 + 960)", "Total Fleet" };
            var realHaulingUnits = new List<string> { "N°", "%", "%", "%", "h", "h", "h", "h", "h", "h", "h", "h", "h", "t/h", "min", "kt" };

            realHaulingWorksheet.Column(1).Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            realHaulingWorksheet.Column(1).Width = 31;
            realHaulingWorksheet.Cells["A1:A2"].Merge = true;
            realHaulingWorksheet.Cells["A1"].Value = "Total Hauling Fleet";
            realHaulingWorksheet.Cells["A1:C1"].Style.Font.Bold = true;
            realHaulingWorksheet.Cells["A1"].Style.Font.Size = 16;
            realHaulingWorksheet.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.BlackFontMineCompliance));
            realHaulingWorksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            realHaulingWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realHaulingWorksheet.Column(2).Width = 16;

            realHaulingWorksheet.Cells["C4"].Style.Font.Bold = true;
            realHaulingWorksheet.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realHaulingWorksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realHaulingWorksheet.Cells["C4"].Value = "Month";
            realHaulingWorksheet.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realHaulingWorksheet.Cells["C4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            realHaulingWorksheet.Cells["C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            realHaulingWorksheet.Cells["C4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            int[] rows8 = { 4, 22, 40, 58, 76, 94, 112, 130, 148, 166, 184, 202, 220 };
            for (var i = rows8.GetLowerBound(0); i <= rows8.GetUpperBound(0); i++)
            {
                realHaulingWorksheet.Cells[rows8[i], 1].Value = realHaulingHeaders[i];
                realHaulingWorksheet.Cells[rows8[i], 1].Style.Font.Bold = true;
                realHaulingWorksheet.Cells[rows8[i], 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                realHaulingWorksheet.Cells[rows8[i], 2].Value = "Unit";
                realHaulingWorksheet.Cells[rows8[i], 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
                realHaulingWorksheet.Cells[rows8[i] + 16, 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

                realHaulingWorksheet.Cells[$"A{rows8[i]}:B{rows8[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                realHaulingWorksheet.Cells[$"A{rows8[i]}:C{rows8[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                realHaulingWorksheet.Cells[$"A{rows8[i]}:B{rows8[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                if (i % 2 != 0)
                {
                    realHaulingWorksheet.Cells[rows8[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i] + 16, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i] + 16, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i] + 16, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i] + 16, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 }:A{rows8[i] + 15 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 }:A{rows8[i] + 15 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[$"B{rows8[i] + 1 }:B{rows8[i] + 15}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[$"B{rows8[i] + 1 }:B{rows8[i] + 15}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    realHaulingWorksheet.Cells[rows8[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i] + 16, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i] + 16, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[rows8[i] + 16, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[rows8[i] + 16, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 }:A{rows8[i] + 15 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 }:A{rows8[i] + 15 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                    realHaulingWorksheet.Cells[$"B{rows8[i] + 1 }:B{rows8[i] + 15}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    realHaulingWorksheet.Cells[$"B{rows8[i] + 1 }:B{rows8[i] + 15}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
                }

                for (var j = 0; j < realHaulingParameters.Count; j++)
                {
                    realHaulingWorksheet.Cells[rows8[i] + 1 + j, 1].Value = realHaulingParameters[j];
                    realHaulingWorksheet.Cells[rows8[i] + 1 + j, 2].Value = realHaulingUnits[j];
                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 + j}:C{rows8[i] + 1 + j}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    realHaulingWorksheet.Cells[$"A{rows8[i] + 1 + j}:C{rows8[i] + 1 + j}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
            }
            realHaulingWorksheet.Column(3).Width = 11;

            var realMillWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.RealMillFCMineComplianceWorksheet);
            var realMillParameters = new List<string> { "Ore Grade - CuT", "Mill Recovery", "Mill Feed" };
            var realMillUnits = new List<string> { "%", "%", "dmt" };

            realMillWorksheet.Cells["A1:A2"].Merge = true;
            realMillWorksheet.Cells["A1"].Value = "Mill";
            realMillWorksheet.Cells["A1"].Style.Font.Bold = true;
            realMillWorksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            realMillWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            realMillWorksheet.Cells["A1"].Style.Font.Size = 16;

            realMillWorksheet.Cells["B3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            realMillWorksheet.Cells["B3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMillWorksheet.Cells["B3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

            realMillWorksheet.Cells["C3"].Style.Font.Bold = true;
            realMillWorksheet.Cells["B3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            realMillWorksheet.Cells["B3"].Value = "Item";
            realMillWorksheet.Cells["C3"].Value = "Month";
            realMillWorksheet.Cells[$"A4:C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            realMillWorksheet.Cells["C3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMillWorksheet.Cells["C3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            realMillWorksheet.Cells["B3:C3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            realMillWorksheet.Cells["A3:C3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < realMillParameters.Count; i++)
            {
                realMillWorksheet.Column(1 + i).Width = 16;
                realMillWorksheet.Cells[4 + i, 1].Value = realMillParameters[i];
                realMillWorksheet.Cells[4 + i, 2].Value = realMillUnits[i];
                realMillWorksheet.Cells[$"A{4 + i}:C{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                realMillWorksheet.Cells[$"A{4 + i}:C{4 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            realMillWorksheet.Cells["A4:A6"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            realMillWorksheet.Cells["A4:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMillWorksheet.Cells["A4:A6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
            realMillWorksheet.Cells["B4:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            realMillWorksheet.Cells["B4:B6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = MineComplianceConstants.RealMineComplianceExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledGenerateRealTemplate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadMineComplianceRealTemplate()
        {
            _realMovementProduction.Clear();
            _realPitDisintegrated.Clear();
            _realMillFc.Clear();
            _realLoadingFc.Clear();
            _realHaulingFc.Clear();

            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var realMovementProductionTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.RealMovementProductionMineComplianceWorksheet];
                var realPitDisintegratedTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.RealPitDisintegratedMineComplianceWorksheet];
                var realMillFCTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.RealMillFCMineComplianceWorksheet];
                var realLoadingFCTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.RealLoadingFCMineComplianceWorksheet];
                var realHaulingFCTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.RealHaulingFCMineComplianceWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - MineComplianceConstants.RealMineComplianceExcelFileName.Length) == MineComplianceConstants.RealMineComplianceExcelFileName)
                {
                    try
                    {
                        var openWriteCheck = File.OpenWrite(openFileDialog.FileName);
                        openWriteCheck.Close();

                        for (var i = 0; i < 45; i++)
                            if (realMovementProductionTemplateWorksheet.Cells[3 + i, 4].Value == null)
                                realMovementProductionTemplateWorksheet.Cells[3 + i, 4].Value = 0;

                        _realMovementProduction.Add(new MineComplianceRealMovementProduction()
                        {
                            LosColoradosOreGradeCutPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[3, 4].Value.ToString())/100,
                            LosColoradosMillRecoveryPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[4, 4].Value.ToString())/100,
                            LosColoradosMillFeedTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[5, 4].Value.ToString())/1000,
                            LosColoradosCuExMillTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[6, 4].Value.ToString())/1000,
                            LosColoradosRuntimePercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[7, 4].Value.ToString())/100,
                            LosColoradosHoursHours = double.Parse(realMovementProductionTemplateWorksheet.Cells[8, 4].Value.ToString()),
                            LagunaSecaOreGradeCutPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[11, 4].Value.ToString())/100,
                            LagunaSecaMillRecoveryPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[12, 4].Value.ToString())/100,
                            LagunaSecaMillFeedTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[13, 4].Value.ToString())/1000,
                            LagunaSecaCuExMillTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[14, 4].Value.ToString())/1000,
                            LagunaSecaRuntimePercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[15, 4].Value.ToString())/100,
                            LagunaSecaHoursHours = double.Parse(realMovementProductionTemplateWorksheet.Cells[16, 4].Value.ToString()),
                            LagunaSeca2OreGradeCutPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[19, 4].Value.ToString())/100,
                            LagunaSeca2MillRecoveryPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[20, 4].Value.ToString())/100,
                            LagunaSeca2MillFeedTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[21, 4].Value.ToString())/1000,
                            LagunaSeca2CuExMillTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[22, 4].Value.ToString())/1000,
                            LagunaSeca2RuntimePercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[23, 4].Value.ToString())/100,
                            LagunaSeca2HoursHours = double.Parse(realMovementProductionTemplateWorksheet.Cells[24, 4].Value.ToString()),
                            OxideCutOreGradeRomPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[27, 4].Value.ToString())/100,
                            OxideCusOreGradeRomPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[28, 4].Value.ToString())/100,
                            OxideCutOreGradeCrusherPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[29, 4].Value.ToString())/100,
                            OxideCusOreGradeCrusherPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[30, 4].Value.ToString())/100,
                            OxideRecoveryCrusherAndRomPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[31, 4].Value.ToString())/100,
                            OxideRomTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[32, 4].Value.ToString())/1000000,
                            OxideCrushedMaterialTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[33, 4].Value.ToString())/1000,
                            OxideTotalStackedMaterialTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[34, 4].Value.ToString())/1000000,
                            OxideCuCathodesTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[35, 4].Value.ToString())/1000,
                            SulphideLeachCutOreGradePercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[36, 4].Value.ToString())/100,
                            SulphideLeachCusOreGradePercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[37, 4].Value.ToString())/100,
                            SulphideLeachRecoveryPercentage = double.Parse(realMovementProductionTemplateWorksheet.Cells[38, 4].Value.ToString())/100,
                            SulphideLeachStackedMaterialFromMineTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[39, 4].Value.ToString())/1000,
                            SulphideLeachContractorsStackedMaterialFromStocksTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[40, 4].Value.ToString())/1000,
                            SulphideLeachMelStackedMaterialFromStocksTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[41, 4].Value.ToString())/1000,
                            SulphideLeachTotalStackedMaterialTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[42, 4].Value.ToString())/1000000,
                            SulphideLeachCuCathodesTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[43, 4].Value.ToString())/1000,
                            ColosoCuExColosoTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[46, 4].Value.ToString())/1000,
                            ColosoCuFromLowGradeConcentrateTonnes = double.Parse(realMovementProductionTemplateWorksheet.Cells[47, 4].Value.ToString())/1000
                        });

                        for (var i = 0; i < 8; i++)
                            if (realPitDisintegratedTemplateWorksheet.Cells[6, 3 + i].Value == null)
                                realPitDisintegratedTemplateWorksheet.Cells[6, 3 + i].Value = 0;

                        _realPitDisintegrated.Add(new MineComplianceRealPitDisintegrated()
                        {
                            ExpitEsTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 3].Value.ToString())/1000000,
                            ExpitEnTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 4].Value.ToString())/1000000,
                            TotalExpitTonnes = (double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 3].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 4].Value.ToString())) / 1000000,

                            MillRehandlingTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 5].Value.ToString()) / 1000000,
                            OlRehandlingTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 6].Value.ToString()) / 1000000,
                            SlRehandlingTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 7].Value.ToString()) / 1000000,
                            OtherRehandlingTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 8].Value.ToString()) / 1000000,
                            //(MillRehandlingTonnes + OlRehandlingTonnes + SlRehandlingTonnes + OtherRehandlingTonnes)/1000000
                            TotalRehandlingTonnes = (double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 5].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 6].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 7].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 8].Value.ToString())) / 1000000,
                            //(ExpitEsTonnes + ExpitEnTonnes + MillRehandlingTonnes + OlRehandlingTonnes + SlRehandlingTonnes + OtherRehandlingTonnes)/1000000
                            TotalMovementTonnes = (double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 3].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 4].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 5].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 6].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 7].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 8].Value.ToString())) / 1000000,

                            RehandlingTotalTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 9].Value.ToString()) / 1000000,
                            MovementTotalTonnes = double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 10].Value.ToString()) / 1000000,
                            //(RehandlingTotalTonnes + MovementTotalTonnes)/1000000
                            TotalTonnes = (double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 9].Value.ToString()) + double.Parse(realPitDisintegratedTemplateWorksheet.Cells[6, 10].Value.ToString())) / 1000000
                        });

                        for (var i = 0; i < 3; i++)
                            if (realMillFCTemplateWorksheet.Cells[4 + i, 3].Value == null)
                                realMillFCTemplateWorksheet.Cells[4 + i, 3].Value = 0;

                        _realMillFc.Add(new MineComplianceRealMillFc()
                        {
                            OreGradeCut = double.Parse(realMillFCTemplateWorksheet.Cells[4, 3].Value.ToString())/100,
                            MillRecovery = double.Parse(realMillFCTemplateWorksheet.Cells[5, 3].Value.ToString())/100,
                            MillFeed = double.Parse(realMillFCTemplateWorksheet.Cells[6, 3].Value.ToString())/1000000
                        });

                        int[] rows1 = { 0, 16, 32, 48, 64, 80, 96, 112, 128 };
                        for (var i = rows1.GetLowerBound(0); i <= rows1.GetUpperBound(0); i++)
                        {
                            for (var j = 1; j < 15; j++)
                            {
                                if (realLoadingFCTemplateWorksheet.Cells[4 + rows1[i] + j, 3].Value == null)
                                    realLoadingFCTemplateWorksheet.Cells[4 + rows1[i] + j, 3].Value = 0;
                            }

                            _realLoadingFc.Add(new MineComplianceRealLoadingFc()
                            {
                                Name = realLoadingFCTemplateWorksheet.Cells[4 + rows1[i], 1].Value.ToString(),
                                Units = double.Parse(realLoadingFCTemplateWorksheet.Cells[5 + rows1[i], 3].Value.ToString()),
                                AvailabilityPercentage = double.Parse(realLoadingFCTemplateWorksheet.Cells[6 + rows1[i], 3].Value.ToString())/100,
                                UtilizationPercentage = double.Parse(realLoadingFCTemplateWorksheet.Cells[7 + rows1[i], 3].Value.ToString())/100,
                                TotalHoursHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[8 + rows1[i], 3].Value.ToString()),
                                AvailableHoursHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[9 + rows1[i], 3].Value.ToString()),
                                EquipmentScheduledDowntimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[10 + rows1[i], 3].Value.ToString()),
                                EquipmentNonScheduledDowntimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[11 + rows1[i], 3].Value.ToString()),
                                ProcessScheduledDowntimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[12 + rows1[i], 3].Value.ToString()),
                                ProcessNonScheduledDowntimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[13 + rows1[i], 3].Value.ToString()),
                                StandByHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[14 + rows1[i], 3].Value.ToString()),
                                HangTimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[15 + rows1[i], 3].Value.ToString()),
                                ProductionTimeHours = double.Parse(realLoadingFCTemplateWorksheet.Cells[16 + rows1[i], 3].Value.ToString()),

                                PerformanceTonnesPerHour = double.Parse(realLoadingFCTemplateWorksheet.Cells[17 + rows1[i], 3].Value.ToString()) / 1000000,
                                TotalTonnesTonnes = double.Parse(realLoadingFCTemplateWorksheet.Cells[18 + rows1[i], 3].Value.ToString()) / 1000
                            });
                        }

                        int[] rows2 = { 0, 18, 36, 54, 72, 90, 108, 126, 144, 162, 180, 198, 216 };

                        for (var i = rows2.GetLowerBound(0); i <= rows2.GetUpperBound(0); i++)
                        {
                            for (var j = 1; j < 21; j++)
                                if (realHaulingFCTemplateWorksheet.Cells[4 + rows2[i] + j, 3].Value == null)
                                    realHaulingFCTemplateWorksheet.Cells[4 + rows2[i] + j, 3].Value = 0;

                            _realHaulingFc.Add(new MineComplianceRealHaulingFc()
                            {
                                Name = realHaulingFCTemplateWorksheet.Cells[4 + rows2[i], 1].Value.ToString(),
                                Units = double.Parse(realHaulingFCTemplateWorksheet.Cells[5 + rows2[i], 3].Value.ToString()),
                                MechanicalAvailabilityPercentage = double.Parse(realHaulingFCTemplateWorksheet.Cells[6 + rows2[i], 3].Value.ToString())/100,
                                PhysicalAvailabilityPercentage = double.Parse(realHaulingFCTemplateWorksheet.Cells[7 + rows2[i], 3].Value.ToString())/100,
                                UtilizationPercentage = double.Parse(realHaulingFCTemplateWorksheet.Cells[8 + rows2[i], 3].Value.ToString())/100,
                                TotalHoursHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[9 + rows2[i], 3].Value.ToString()),
                                AvailableHoursHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[10 + rows2[i], 3].Value.ToString()),
                                EquipmentScheduledDowntimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[11 + rows2[i], 3].Value.ToString()),
                                EquipmentNonScheduledDowntimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[12 + rows2[i], 3].Value.ToString()),
                                ProcessScheduledDowntimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[13 + rows2[i], 3].Value.ToString()),
                                ProcessNonScheduledDowntimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[14 + rows2[i], 3].Value.ToString()),
                                StandByHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[15 + rows2[i], 3].Value.ToString()),
                                QueueTimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[16 + rows2[i], 3].Value.ToString()),
                                ProductionTimeHoursHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[17 + rows2[i], 3].Value.ToString()),
                                PerformanceTonnesPerHour = double.Parse(realHaulingFCTemplateWorksheet.Cells[18 + rows2[i], 3].Value.ToString()) / 1000000,
                                CycleTimeHours = double.Parse(realHaulingFCTemplateWorksheet.Cells[19 + rows2[i], 3].Value.ToString()) / 60,
                                TotalTonnesTonnes = double.Parse(realHaulingFCTemplateWorksheet.Cells[20 + rows2[i], 3].Value.ToString()) / 1000,
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

                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineComplianceExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);
                
                if (loadFileInfo.Exists)
                {
                    var package = new ExcelPackage(loadFileInfo);
                    var realMovementProductionWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealMovementProductionMineComplianceSpotfireWorksheet];
                    var realPitDisintegratedWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealPitDisintegratedMineComplianceSpotfireWorksheet];
                    var realMillWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealMillFCMineComplianceSpotfireWorksheet];
                    var realLoadingWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealLoadingFCMineComplianceSpotfireWorksheet];
                    var realLoadingSSWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealLoadingSSYTDMineComplianceSpotfireWorksheet];
                    var realHaulingWorksheet = package.Workbook.Worksheets[MineComplianceConstants.RealHaulingFCMineComplianceSpotfireWorksheet];
                    var realHaulingTF = package.Workbook.Worksheets[MineComplianceConstants.RealHaulingTFYTDMineComplianceSpotfireWorksheet];

                    if (realMovementProductionWorksheet != null & realPitDisintegratedWorksheet != null & realMillWorksheet != null & realLoadingWorksheet != null & realHaulingWorksheet != null)
                    {
                        try
                        {
                            var openWriteCheck = File.OpenWrite(loadFilePath);
                            openWriteCheck.Close();

                            var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                            var lastRow1 = realMovementProductionWorksheet.Dimension.End.Row + 1;

                            realMovementProductionWorksheet.Cells[lastRow1, 1].Value = newDate;
                            realMovementProductionWorksheet.Cells[lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            realMovementProductionWorksheet.Cells[lastRow1, 2].Value = _realMovementProduction[0].LosColoradosOreGradeCutPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 3].Value = _realMovementProduction[0].LosColoradosMillRecoveryPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 4].Value = _realMovementProduction[0].LosColoradosMillFeedTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 5].Value = _realMovementProduction[0].LosColoradosCuExMillTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 6].Value = _realMovementProduction[0].LosColoradosRuntimePercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 7].Value = _realMovementProduction[0].LosColoradosHoursHours;
                            realMovementProductionWorksheet.Cells[lastRow1, 8].Value = _realMovementProduction[0].LagunaSecaOreGradeCutPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 9].Value = _realMovementProduction[0].LagunaSecaMillRecoveryPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 10].Value = _realMovementProduction[0].LagunaSecaMillFeedTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 11].Value = _realMovementProduction[0].LagunaSecaCuExMillTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 12].Value = _realMovementProduction[0].LagunaSecaRuntimePercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 13].Value = _realMovementProduction[0].LagunaSecaHoursHours;
                            realMovementProductionWorksheet.Cells[lastRow1, 14].Value = _realMovementProduction[0].LagunaSeca2OreGradeCutPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 15].Value = _realMovementProduction[0].LagunaSeca2MillRecoveryPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 16].Value = _realMovementProduction[0].LagunaSeca2MillFeedTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 17].Value = _realMovementProduction[0].LagunaSeca2CuExMillTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 18].Value = _realMovementProduction[0].LagunaSeca2RuntimePercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 19].Value = _realMovementProduction[0].LagunaSeca2HoursHours;
                            realMovementProductionWorksheet.Cells[lastRow1, 20].Value = _realMovementProduction[0].OxideCutOreGradeRomPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 21].Value = _realMovementProduction[0].OxideCusOreGradeRomPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 22].Value = _realMovementProduction[0].OxideCutOreGradeCrusherPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 23].Value = _realMovementProduction[0].OxideCusOreGradeCrusherPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 24].Value = _realMovementProduction[0].OxideRecoveryCrusherAndRomPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 25].Value = _realMovementProduction[0].OxideRomTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 26].Value = _realMovementProduction[0].OxideCrushedMaterialTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 27].Value = _realMovementProduction[0].OxideTotalStackedMaterialTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 28].Value = _realMovementProduction[0].OxideCuCathodesTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 29].Value = _realMovementProduction[0].SulphideLeachCutOreGradePercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 30].Value = _realMovementProduction[0].SulphideLeachCusOreGradePercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 31].Value = _realMovementProduction[0].SulphideLeachRecoveryPercentage;
                            realMovementProductionWorksheet.Cells[lastRow1, 32].Value = _realMovementProduction[0].SulphideLeachStackedMaterialFromMineTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 33].Value = _realMovementProduction[0].SulphideLeachContractorsStackedMaterialFromStocksTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 34].Value = _realMovementProduction[0].SulphideLeachMelStackedMaterialFromStocksTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 35].Value = _realMovementProduction[0].SulphideLeachTotalStackedMaterialTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 36].Value = _realMovementProduction[0].SulphideLeachCuCathodesTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 37].Value = _realMovementProduction[0].ColosoCuExColosoTonnes;
                            realMovementProductionWorksheet.Cells[lastRow1, 38].Value = _realMovementProduction[0].ColosoCuFromLowGradeConcentrateTonnes;

                            var lastRow2 = realPitDisintegratedWorksheet.Dimension.End.Row + 1;

                            realPitDisintegratedWorksheet.Cells[lastRow2, 1].Value = newDate;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            realPitDisintegratedWorksheet.Cells[lastRow2, 2].Value = _realPitDisintegrated[0].ExpitEsTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 3].Value = _realPitDisintegrated[0].ExpitEnTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 4].Value = _realPitDisintegrated[0].TotalExpitTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 5].Value = _realPitDisintegrated[0].MillRehandlingTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 6].Value = _realPitDisintegrated[0].OlRehandlingTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 7].Value = _realPitDisintegrated[0].SlRehandlingTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 8].Value = _realPitDisintegrated[0].OtherRehandlingTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 9].Value = _realPitDisintegrated[0].TotalRehandlingTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 10].Value = _realPitDisintegrated[0].TotalMovementTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 11].Value = _realPitDisintegrated[0].RehandlingTotalTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 12].Value = _realPitDisintegrated[0].MovementTotalTonnes;
                            realPitDisintegratedWorksheet.Cells[lastRow2, 13].Value = _realPitDisintegrated[0].TotalTonnes;

                            var lastRow3 = realMillWorksheet.Dimension.End.Row + 1;

                            realMillWorksheet.Cells[lastRow3, 1].Value = newDate;
                            realMillWorksheet.Cells[lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            realMillWorksheet.Cells[lastRow3, 2].Value = _realMillFc[0].OreGradeCut;
                            realMillWorksheet.Cells[lastRow3, 3].Value = _realMillFc[0].MillRecovery;
                            realMillWorksheet.Cells[lastRow3, 4].Value = _realMillFc[0].MillFeed;

                            var lastRow4 = realLoadingWorksheet.Dimension.End.Row + 1;
                            var lastRow5 = realLoadingSSWorksheet.Dimension.End.Row + 1;

                            for (var i = 0; i < _realLoadingFc.Count; i++)
                            {
                                realLoadingWorksheet.Cells[i + lastRow4, 1].Value = newDate;
                                realLoadingWorksheet.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                realLoadingWorksheet.Cells[i + lastRow4, 2].Value = _realLoadingFc[i].Name;
                                realLoadingWorksheet.Cells[i + lastRow4, 3].Value = _realLoadingFc[i].Units;
                                realLoadingWorksheet.Cells[i + lastRow4, 4].Value = _realLoadingFc[i].AvailabilityPercentage;
                                realLoadingWorksheet.Cells[i + lastRow4, 5].Value = _realLoadingFc[i].UtilizationPercentage;
                                realLoadingWorksheet.Cells[i + lastRow4, 6].Value = _realLoadingFc[i].TotalHoursHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 7].Value = _realLoadingFc[i].AvailableHoursHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 8].Value = _realLoadingFc[i].EquipmentScheduledDowntimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 9].Value = _realLoadingFc[i].EquipmentNonScheduledDowntimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 10].Value = _realLoadingFc[i].ProcessScheduledDowntimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 11].Value = _realLoadingFc[i].ProcessNonScheduledDowntimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 12].Value = _realLoadingFc[i].StandByHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 13].Value = _realLoadingFc[i].HangTimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 14].Value = _realLoadingFc[i].ProductionTimeHours;
                                realLoadingWorksheet.Cells[i + lastRow4, 15].Value = _realLoadingFc[i].PerformanceTonnesPerHour;
                                realLoadingWorksheet.Cells[i + lastRow4, 16].Value = _realLoadingFc[i].TotalTonnesTonnes;

                                if (_realLoadingFc[i].Name == "SUMMARY SHOVEL 73 yd3")
                                {
                                    realLoadingSSWorksheet.Cells[lastRow5, 1].Value = newDate;
                                    realLoadingSSWorksheet.Cells[lastRow5, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    realLoadingSSWorksheet.Cells[lastRow5, 2].Value = _realLoadingFc[i].Name;
                                    realLoadingSSWorksheet.Cells[lastRow5, 3].Value = _realLoadingFc[i].Units;
                                    realLoadingSSWorksheet.Cells[lastRow5, 4].Value = _realLoadingFc[i].AvailabilityPercentage;
                                    realLoadingSSWorksheet.Cells[lastRow5, 5].Value = _realLoadingFc[i].UtilizationPercentage;
                                    realLoadingSSWorksheet.Cells[lastRow5, 6].Value = _realLoadingFc[i].TotalHoursHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 7].Value = _realLoadingFc[i].AvailableHoursHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 8].Value = _realLoadingFc[i].EquipmentScheduledDowntimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 9].Value = _realLoadingFc[i].EquipmentNonScheduledDowntimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 10].Value = _realLoadingFc[i].ProcessScheduledDowntimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 11].Value = _realLoadingFc[i].ProcessNonScheduledDowntimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 12].Value = _realLoadingFc[i].StandByHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 13].Value = _realLoadingFc[i].HangTimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 14].Value = _realLoadingFc[i].ProductionTimeHours;
                                    realLoadingSSWorksheet.Cells[lastRow5, 15].Value = _realLoadingFc[i].PerformanceTonnesPerHour;
                                    realLoadingSSWorksheet.Cells[lastRow5, 16].Value = _realLoadingFc[i].TotalTonnesTonnes;
                                }

                                if (_realLoadingFc[i].Name == "TOTAL FLEET")
                                {
                                    realLoadingSSWorksheet.Cells[lastRow5, 17].Value = _realLoadingFc[i].TotalTonnesTonnes;
                                }
                            }

                            var lastRow6 = realHaulingWorksheet.Dimension.End.Row + 1;
                            var lastRow7 = realHaulingTF.Dimension.End.Row + 1;

                            for (var i = 0; i < _realHaulingFc.Count; i++)
                            {
                                realHaulingWorksheet.Cells[i + lastRow6, 1].Value = newDate;
                                realHaulingWorksheet.Cells[i + lastRow6, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                realHaulingWorksheet.Cells[i + lastRow6, 2].Value = _realHaulingFc[i].Name;
                                realHaulingWorksheet.Cells[i + lastRow6, 3].Value = _realHaulingFc[i].Units;
                                realHaulingWorksheet.Cells[i + lastRow6, 4].Value = _realHaulingFc[i].MechanicalAvailabilityPercentage;
                                realHaulingWorksheet.Cells[i + lastRow6, 5].Value = _realHaulingFc[i].PhysicalAvailabilityPercentage;
                                realHaulingWorksheet.Cells[i + lastRow6, 6].Value = _realHaulingFc[i].UtilizationPercentage;
                                realHaulingWorksheet.Cells[i + lastRow6, 7].Value = _realHaulingFc[i].TotalHoursHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 8].Value = _realHaulingFc[i].AvailableHoursHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 9].Value = _realHaulingFc[i].EquipmentScheduledDowntimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 10].Value = _realHaulingFc[i].EquipmentNonScheduledDowntimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 11].Value = _realHaulingFc[i].ProcessScheduledDowntimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 12].Value = _realHaulingFc[i].ProcessNonScheduledDowntimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 13].Value = _realHaulingFc[i].StandByHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 14].Value = _realHaulingFc[i].QueueTimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 15].Value = _realHaulingFc[i].ProductionTimeHoursHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 16].Value = _realHaulingFc[i].PerformanceTonnesPerHour;
                                realHaulingWorksheet.Cells[i + lastRow6, 17].Value = _realHaulingFc[i].CycleTimeHours;
                                realHaulingWorksheet.Cells[i + lastRow6, 18].Value = _realHaulingFc[i].TotalTonnesTonnes;

                                if (_realHaulingFc[i].Name == "Total Fleet")
                                {
                                    realHaulingTF.Cells[lastRow7, 1].Value = newDate;
                                    realHaulingTF.Cells[lastRow7, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    realHaulingTF.Cells[lastRow7, 2].Value = _realHaulingFc[i].Name;
                                    realHaulingTF.Cells[lastRow7, 3].Value = _realHaulingFc[i].Units;
                                    realHaulingTF.Cells[lastRow7, 4].Value = _realHaulingFc[i].MechanicalAvailabilityPercentage;
                                    realHaulingTF.Cells[lastRow7, 5].Value = _realHaulingFc[i].PhysicalAvailabilityPercentage;
                                    realHaulingTF.Cells[lastRow7, 6].Value = _realHaulingFc[i].UtilizationPercentage;
                                    realHaulingTF.Cells[lastRow7, 7].Value = _realHaulingFc[i].TotalHoursHours;
                                    realHaulingTF.Cells[lastRow7, 8].Value = _realHaulingFc[i].AvailableHoursHours;
                                    realHaulingTF.Cells[lastRow7, 9].Value = _realHaulingFc[i].EquipmentScheduledDowntimeHours;
                                    realHaulingTF.Cells[lastRow7, 10].Value = _realHaulingFc[i].EquipmentNonScheduledDowntimeHours;
                                    realHaulingTF.Cells[lastRow7, 11].Value = _realHaulingFc[i].ProcessScheduledDowntimeHours;
                                    realHaulingTF.Cells[lastRow7, 12].Value = _realHaulingFc[i].ProcessNonScheduledDowntimeHours;
                                    realHaulingTF.Cells[lastRow7, 13].Value = _realHaulingFc[i].StandByHours;
                                    realHaulingTF.Cells[lastRow7, 14].Value = _realHaulingFc[i].QueueTimeHours;
                                    realHaulingTF.Cells[lastRow7, 15].Value = _realHaulingFc[i].ProductionTimeHoursHours;
                                    realHaulingTF.Cells[lastRow7, 16].Value = _realHaulingFc[i].PerformanceTonnesPerHour;
                                    realHaulingTF.Cells[lastRow7, 17].Value = _realHaulingFc[i].CycleTimeHours;
                                    realHaulingTF.Cells[lastRow7, 18].Value = _realHaulingFc[i].TotalTonnesTonnes;
                                }
                            }
                            byte[] fileText2 = package.GetAsByteArray();
                            File.WriteAllBytes(loadFilePath, fileText2);
                            MyLastDateRefreshRealValues = $"{StringResources.Updated}: {DateTime.Now}";
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

        private void GenerateMineComplianceBudgetTemplate()
        {
            var months = new List<string> { "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            var budgetPitDisintegratedCategories = new List<string> { "Category", "Parameter", "Unit", "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            var budgetPitDisintegratedParameters = new List<string> { "Expit ES", "Expit EN", "Mill", "OL", "SL", "Other", "Reh.", "Mov." };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = MineComplianceConstants.BudgetMineComplianceWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";

            var budgetPitDisintegratedWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.BudgetPitDisintegratedMineComplianceWorksheet);
            var budgetPrincipalWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.BudgetPrincipalMineComplianceWorksheet);
            var budgetMovementProductionWorksheet = excelPackage.Workbook.Worksheets.Add(MineComplianceConstants.BudgetMovementProductionMineComplianceWorksheet);

            budgetPitDisintegratedWorksheet.Cells["B2"].Value = $"FY{MyFiscalYear}";
            budgetPitDisintegratedWorksheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["B2:B5"].Style.Font.Bold = true;
            budgetPitDisintegratedWorksheet.Cells["B3:B5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["B3:B5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGrayBackgroundMineCompliance));

            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Merge = true;
            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Style.Font.Bold = true;
            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Value = "Budget";
            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["B2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["C2:J2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.BlueBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells[$"B2:J2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < 16; i++)           
                budgetPitDisintegratedWorksheet.Cells[$"B{2 + i}:J{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            string[] columns1 = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
            for (var i = columns1.GetLowerBound(0); i <= columns1.GetUpperBound(0); i++)
            {
                budgetPitDisintegratedWorksheet.Cells[$"{columns1[i]}2:{columns1[i]}17"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                budgetPitDisintegratedWorksheet.Column(2 + i).Width = 14;
            }

            budgetPitDisintegratedWorksheet.Column(1).Width = 3;
            budgetPitDisintegratedWorksheet.Cells["A6"].Value = "Month";
            budgetPitDisintegratedWorksheet.Cells["A6:A17"].Merge = true;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.Font.Bold = true;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.TextRotation = 90;
            budgetPitDisintegratedWorksheet.Cells["A6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            budgetPitDisintegratedWorksheet.Cells["A17"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            for (var i = 0; i < budgetPitDisintegratedCategories.Count; i++)            
                budgetPitDisintegratedWorksheet.Cells[3 + i, 2].Value = budgetPitDisintegratedCategories[i];

            budgetPitDisintegratedWorksheet.Cells["C3:D3"].Merge = true;
            budgetPitDisintegratedWorksheet.Cells["E3:H3"].Merge = true;
            budgetPitDisintegratedWorksheet.Cells["I3:J3"].Merge = true;
            budgetPitDisintegratedWorksheet.Cells["C3:J3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
            budgetPitDisintegratedWorksheet.Cells[3, 3].Value = "Movement";
            budgetPitDisintegratedWorksheet.Cells[3, 5].Value = "Rehandling";
            budgetPitDisintegratedWorksheet.Cells[3, 9].Value = "Total";

            budgetPitDisintegratedWorksheet.Cells["C3:J3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["C3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["E3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["E3:H3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));

            for (var i = 0; i < 3; i++)            
                budgetPitDisintegratedWorksheet.Cells[$"C{3 + i}:J{3 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < budgetPitDisintegratedParameters.Count; i++)
            {
                budgetPitDisintegratedWorksheet.Cells[4, 3 + i].Value = budgetPitDisintegratedParameters[i];
                budgetPitDisintegratedWorksheet.Cells[5, 3 + i].Value = "t";
            }
            budgetPitDisintegratedWorksheet.Cells["C4:J4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            budgetPitDisintegratedWorksheet.Cells["C4:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["C4:J4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["E4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["E4:H4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

            budgetPitDisintegratedWorksheet.Cells["C5:J5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["C5:J5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
            budgetPitDisintegratedWorksheet.Cells["E5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPitDisintegratedWorksheet.Cells["E5:H5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
     
            var budgetMovementProductionZones = new List<string> { "Los Colorados", "Laguna Seca", "Laguna Seca 2", "Oxide", "Sulphide Leach" };
            var budgetMovementProductionParameters1 = new List<string> { "Ore Grade - CuT", "Mill Recovery", "Mill Feed", "Cu Ex Mill", "Runtime", "Hours" };
            var budgetMovementProdctionParameters2 = new List<string> { "Ore to OL", "Cu Cathodes", "Stacked Material from Mine", "Contractors Stacked Material from Stocks", "MEL Stacked Material from Stocks", "Total Stacked Material", "Cu Cathodes" };
            var budgetMovementProductionUnits1 = new List<string> { "%", "%", "t", "t Cu", "%", "h" };
            var budgetMovementProductionUnits2 = new List<string> { "dmt", "t Cu", "dmt", "dmt", "dmt", "dmt", "t Cu" };

            budgetMovementProductionWorksheet.Column(1).Width = 20;
            budgetMovementProductionWorksheet.Cells["A1"].Value = "Budget";
            budgetMovementProductionWorksheet.Row(1).Style.Font.Bold = true;
            budgetMovementProductionWorksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetMovementProductionWorksheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetMovementProductionWorksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkBlueBackgroundMineCompliance));
            budgetMovementProductionWorksheet.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            budgetMovementProductionWorksheet.Cells["D1"].Value = $"FY{MyFiscalYear}";
            budgetMovementProductionWorksheet.Cells["D1:O1"].Merge = true;

            budgetMovementProductionWorksheet.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            budgetMovementProductionWorksheet.Cells["D1:O1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetMovementProductionWorksheet.Cells["D1:O1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGrayBackgroundMineCompliance));
            budgetMovementProductionWorksheet.Cells["D1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            budgetMovementProductionWorksheet.Cells["O1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            budgetMovementProductionWorksheet.Cells["D1:O1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            budgetMovementProductionWorksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            for (var i = 0; i < months.Count; i++)
            {
                budgetMovementProductionWorksheet.Cells[2, 4 + i].Value = months[i];
                budgetMovementProductionWorksheet.Column(4 + i).Width = 11;
            }
            budgetMovementProductionWorksheet.Cells["D2:O2"].Style.Font.Bold = true;
            budgetMovementProductionWorksheet.Cells["D2:O2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            budgetMovementProductionWorksheet.Cells["D2:O2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetMovementProductionWorksheet.Cells["D2:O2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));

            int[] rows1 = { 3, 11, 19, 27, 29 };
            int[] rows2 = { 8, 16, 24, 28, 33 };
            for (var i = rows1.GetLowerBound(0); i <= rows1.GetUpperBound(0); i++)
            {
                budgetMovementProductionWorksheet.Cells[rows1[i], 1].Value = budgetMovementProductionZones[i];
                budgetMovementProductionWorksheet.Cells[$"A{rows1[i]}:O{rows1[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
                budgetMovementProductionWorksheet.Cells[$"B{rows1[i]}:B{rows2[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

                if (i % 2 != 0)
                {
                    budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkOrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    budgetMovementProductionWorksheet.Cells[rows1[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGreenBackgroundMineCompliance));
                }
            }

            budgetMovementProductionWorksheet.Column(2).Width = 35;
            for (var i = 0; i < budgetMovementProductionParameters1.Count; i++)
            {
                budgetMovementProductionWorksheet.Cells[3 + i, 2].Value = budgetMovementProductionParameters1[i];
                budgetMovementProductionWorksheet.Cells[3 + i, 3].Value = budgetMovementProductionUnits1[i];
                budgetMovementProductionWorksheet.Cells[$"B{3 + i}:O{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                budgetMovementProductionWorksheet.Cells[11 + i, 2].Value = budgetMovementProductionParameters1[i];
                budgetMovementProductionWorksheet.Cells[11 + i, 3].Value = budgetMovementProductionUnits1[i];
                budgetMovementProductionWorksheet.Cells[$"B{11 + i}:O{11 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                budgetMovementProductionWorksheet.Cells[19 + i, 2].Value = budgetMovementProductionParameters1[i];
                budgetMovementProductionWorksheet.Cells[19 + i, 3].Value = budgetMovementProductionUnits1[i];
                budgetMovementProductionWorksheet.Cells[$"B{19 + i}:O{19 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] columns2 = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O" };
            for (var i = columns2.GetLowerBound(0); i <= columns2.GetUpperBound(0); i++)
            {
                budgetMovementProductionWorksheet.Cells[$"{columns2[i]}3:{columns2[i]}8"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[$"{columns2[i]}11:{columns2[i]}16"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[$"{columns2[i]}19:{columns2[i]}24"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[$"{columns2[i]}27:{columns2[i]}33"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            int[] rows3 = { 3, 19, 29 };
            int[] rows4 = { 8, 24, 33 };
            for (var i = rows3.GetLowerBound(0); i <= rows3.GetUpperBound(0); i++)
            {
                budgetMovementProductionWorksheet.Cells[$"B{rows3[i]}:B{rows4[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetMovementProductionWorksheet.Cells[$"B{rows3[i]}:B{rows4[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                budgetMovementProductionWorksheet.Cells[$"C{rows3[i]}:C{rows4[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetMovementProductionWorksheet.Cells[$"C{rows3[i]}:C{rows4[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
            }

            int[] rows5 = { 11, 27 };
            int[] rows6 = { 16, 28 };
            for (var i = rows5.GetLowerBound(0); i <= rows5.GetUpperBound(0); i++)
            {
                budgetMovementProductionWorksheet.Cells[$"B{rows5[i]}:B{rows6[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetMovementProductionWorksheet.Cells[$"B{rows5[i]}:B{rows6[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                budgetMovementProductionWorksheet.Cells[$"C{rows5[i]}:C{rows6[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetMovementProductionWorksheet.Cells[$"C{rows5[i]}:C{rows6[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
            }

            budgetMovementProductionWorksheet.Column(3).Width = 13;
            budgetMovementProductionWorksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int[] rows7 = { 2, 10, 18, 26 };
            for (var i = rows7.GetLowerBound(0); i <= rows7.GetUpperBound(0); i++)
            {
                budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));
                budgetMovementProductionWorksheet.Cells[rows7[i], 3].Value = "Unit";
                budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                if (i % 2 != 0)
                {
                    budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));
                }
                else if (i % 2 == 0)
                {
                    budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    budgetMovementProductionWorksheet.Cells[rows7[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));
                }
            }

            for (var i = 0; i < budgetMovementProductionUnits2.Count; i++)
            {
                budgetMovementProductionWorksheet.Cells[27 + i, 2].Value = budgetMovementProdctionParameters2[i];
                budgetMovementProductionWorksheet.Cells[27 + i, 3].Value = budgetMovementProductionUnits2[i];

                budgetMovementProductionWorksheet.Cells[$"B{27 + i}:O{27 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            
            var budgetPrincipalParameters = new List<string> { "Expit", "Rehandling", "Movement", "Shovels units 73 yd3", "Shovels availability", "Shovels utilization", "Shovels performance", "Stand By", "Production time", "Available hours", "Shovel hours", "Trucks units", "Trucks availability", "Trucks utilization", "Trucks performance", "Stand By", "Trucks hours", "Production Time", "Available hours", "Mill Throughput", "Mill Grade", "Mill Recovery", "Mill Rehandle", "OL Throughput", "OL Grade", "Recovery", "CuS", "SL Throughput", "SL Grade", "Recovery", "CuS", "Mill Production", "Cathodes", "Total Production" };
            var budgetPrincipalUnits = new List<string> { "Mt", "Mt", "Mt", "eq", "%", "%", "t/h", "h", "h", "h", "h", "eq", "%", "%", "t/h", "h", "h", "h", "h", "Mt", "Cu %", "%", "%", "Mt", "Cu %", "%", "%", "Mt", "Cu %", "%", "%", "kt", "kt", "kt" };

            budgetPrincipalWorksheet.Column(2).Width = 20;
            budgetPrincipalWorksheet.Column(2).Style.Font.Bold = true;
            budgetPrincipalWorksheet.Cells["B2"].Value = "BUDGET B01";
            budgetPrincipalWorksheet.Row(2).Style.Font.Bold = true;
            budgetPrincipalWorksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetPrincipalWorksheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPrincipalWorksheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.BlueBackgroundMineCompliance));
            budgetPrincipalWorksheet.Cells["B2:B37"].Style.Font.Color.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.WhiteFontMineCompliance));

            budgetPrincipalWorksheet.Cells["D2"].Value = $"FY{MyFiscalYear}";
            budgetPrincipalWorksheet.Cells["D2:O2"].Merge = true;

            budgetPrincipalWorksheet.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetPrincipalWorksheet.Cells["D2:O2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPrincipalWorksheet.Cells["D2:O2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.DarkGrayBackgroundMineCompliance));
            budgetPrincipalWorksheet.Cells["D2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            budgetPrincipalWorksheet.Cells["C3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            budgetPrincipalWorksheet.Cells["O2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            budgetPrincipalWorksheet.Cells["D2:O2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            budgetPrincipalWorksheet.Cells["D2:O2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            budgetPrincipalWorksheet.Cells["B3:O3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            budgetPrincipalWorksheet.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            budgetPrincipalWorksheet.Row(3).Style.Font.Bold = true;

            for (var i = 0; i < months.Count; i++)
            {
                budgetPrincipalWorksheet.Cells[3, 4 + i].Value = months[i];
                budgetPrincipalWorksheet.Column(4 + i).Width = 11;
            }
            budgetPrincipalWorksheet.Cells["D3:O3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            budgetPrincipalWorksheet.Cells["D3:O3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            budgetPrincipalWorksheet.Cells["D3:O3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GrayBackgroundMineCompliance));

            budgetPrincipalWorksheet.Column(3).Width = 10;
            budgetPrincipalWorksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (var i = 0; i < budgetPrincipalParameters.Count; i++)
            {
                budgetPrincipalWorksheet.Cells[4 + i, 2].Value = budgetPrincipalParameters[i];
                budgetPrincipalWorksheet.Cells[4 + i, 3].Value = budgetPrincipalUnits[i];
                budgetPrincipalWorksheet.Cells[$"B{4 + i}:O{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] columns3 = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O" };
            for (var i = columns3.GetLowerBound(0); i <= columns3.GetUpperBound(0); i++)
                budgetPrincipalWorksheet.Cells[$"{columns2[i]}4:{columns2[i]}37"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            
            int[] rows8 = { 4, 15, 27, 35 };
            int[] rows9 = { 6, 22, 30, 37 };
            for (var i = rows8.GetLowerBound(0); i <= rows8.GetUpperBound(0); i++)
            {
                budgetPrincipalWorksheet.Cells[$"B{rows8[i]}:B{rows9[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetPrincipalWorksheet.Cells[$"B{rows8[i]}:B{rows9[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.GreenBackgroundMineCompliance));

                budgetPrincipalWorksheet.Cells[$"C{rows8[i]}:C{rows9[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetPrincipalWorksheet.Cells[$"C{rows8[i]}:C{rows9[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightGreenBackgroundMineCompliance));
                budgetPrincipalWorksheet.Cells[$"D{rows9[i]}:O{rows9[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            }

            int[] rows10 = { 7, 23, 31 };
            int[] rows11 = { 14, 26, 34 };
            for (var i = rows10.GetLowerBound(0); i <= rows10.GetUpperBound(0); i++)
            {
                budgetPrincipalWorksheet.Cells[$"B{rows10[i]}:B{rows11[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetPrincipalWorksheet.Cells[$"B{rows10[i]}:B{rows11[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.OrangeBackgroundMineCompliance));

                budgetPrincipalWorksheet.Cells[$"C{rows10[i]}:C{rows11[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                budgetPrincipalWorksheet.Cells[$"C{rows10[i]}:C{rows11[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(MineComplianceTemplateColors.LightOrangeBackgroundMineCompliance));
                budgetPrincipalWorksheet.Cells[$"D{rows11[i]}:O{rows11[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = MineComplianceConstants.BudgetMineComplianceExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                var fileStream = File.OpenWrite(dialog.FileName);
                fileStream.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledGenerateBudgetTemplate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadMineComplianceBudgetTemplate()
        {
            _budgetPrincipal.Clear();
            _budgetMovementProduction.Clear();
            _budgetPitDisintegrated.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var budgetPrincipalTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.BudgetPrincipalMineComplianceWorksheet];
                var budgetMovementProductionTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.BudgetMovementProductionMineComplianceWorksheet];
                var budgetPitDisintegratedTemplateWorksheet = excelPackage.Workbook.Worksheets[MineComplianceConstants.BudgetPitDisintegratedMineComplianceWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - MineComplianceConstants.BudgetMineComplianceExcelFileName.Length) == MineComplianceConstants.BudgetMineComplianceExcelFileName)
                {
                    try
                    {
                        var fileStream = File.OpenWrite(openFileDialog.FileName);
                        fileStream.Close();
                        
                        var _date = DateTime.Now;

                        for (var i = 0; i < 12; i++)
                        {
                            var _month = DateTime.ParseExact(budgetPrincipalTemplateWorksheet.Cells[3, 4 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;

                            _date = TemplateDates.ConvertDateToFiscalYearDate(i, MyFiscalYear, _month);

                            for (var j = 0; j < 34; j++)
                            {
                                if (budgetPrincipalTemplateWorksheet.Cells[4 + j, 4 + i].Value == null)
                                    budgetPrincipalTemplateWorksheet.Cells[4 + j, 4 + i].Value = 0;

                                if (budgetMovementProductionTemplateWorksheet.Cells[3 + j, 4 + i].Value == null)
                                    budgetMovementProductionTemplateWorksheet.Cells[3 + j, 4 + i].Value = 0;
                            }

                            for (var j = 0; j < 8; j++)
                                if (budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 3 + j].Value == null)
                                    budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 3 + j].Value = 0;

                            _budgetPrincipal.Add(new MineComplianceBudgetPrincipal()
                            {
                                Date = _date,
                                ExpitTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[4, 4 + i].Value.ToString()),
                                RehandlingTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[5, 4 + i].Value.ToString()),
                                MovementTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[6, 4 + i].Value.ToString()),
                                ShovelsUnits73Yd3 = double.Parse(budgetPrincipalTemplateWorksheet.Cells[7, 4 + i].Value.ToString()),
                                ShovelsAvailabilityPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[8, 4 + i].Value.ToString())/100,
                                ShovelsUtilizationPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[9, 4 + i].Value.ToString())/100,
                                ShovelsPerformanceTonnesPerHour = double.Parse(budgetPrincipalTemplateWorksheet.Cells[10, 4 + i].Value.ToString()) / 1000000,
                                ShovelsStandByHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[11, 4 + i].Value.ToString()),
                                ShovelsProductionTimeHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[12, 4 + i].Value.ToString()),
                                ShovelAvailableHoursHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[13, 4 + i].Value.ToString()),
                                ShovelHoursHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[14, 4 + i].Value.ToString()),
                                TrucksUnits = double.Parse(budgetPrincipalTemplateWorksheet.Cells[15, 4 + i].Value.ToString()),
                                TrucksAvailabilityPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[16, 4 + i].Value.ToString())/100,
                                TrucksUtilizationPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[17, 4 + i].Value.ToString())/100,
                                TrucksPerformanceTonnesPerDay = double.Parse(budgetPrincipalTemplateWorksheet.Cells[18, 4 + i].Value.ToString()) / 1000000,
                                TrucksStandByHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[19, 4 + i].Value.ToString()),
                                TrucksHoursHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[20, 4 + i].Value.ToString()),
                                TrucksProductionTimeHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[21, 4 + i].Value.ToString()),
                                TrucksAvailableHoursHours = double.Parse(budgetPrincipalTemplateWorksheet.Cells[22, 4 + i].Value.ToString()),
                                MillThroughputTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[23, 4 + i].Value.ToString()),
                                MillGradeCuPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[24, 4 + i].Value.ToString())/100,
                                MillRecoveryPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[25, 4 + i].Value.ToString())/100,
                                MillRehandlingPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[26, 4 + i].Value.ToString())/100,
                                OlThroughputTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[27, 4 + i].Value.ToString()),
                                OlGradeCuPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[28, 4 + i].Value.ToString())/100,
                                OlRecoveryPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[29, 4 + i].Value.ToString())/100,
                                OlCuSPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[30, 4 + i].Value.ToString())/100,
                                SlThroughputTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[31, 4 + i].Value.ToString()),
                                SlGradeCuPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[32, 4 + i].Value.ToString())/100,
                                SlRecoveryPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[33, 4 + i].Value.ToString())/100,
                                SlCuSPercentage = double.Parse(budgetPrincipalTemplateWorksheet.Cells[34, 4 + i].Value.ToString())/100,
                                MillProductionTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[35, 4 + i].Value.ToString()),
                                CathodesTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[36, 4 + i].Value.ToString()),
                                TotalProductionTonnes = double.Parse(budgetPrincipalTemplateWorksheet.Cells[37, 4 + i].Value.ToString())
                            });

                            _budgetMovementProduction.Add(new MineComplianceBudgetMovementProduction()
                            {
                                Date = _date,
                                LosColoradosOreGradeCutPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[3, 4 + i].Value.ToString())/100,
                                LosColoradosMillRecoveryPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[4, 4 + i].Value.ToString())/100,
                                LosColoradosMillFeedTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[5, 4 + i].Value.ToString()) / 1000,
                                LosColoradosCuExMillTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[6, 4 + i].Value.ToString()) / 1000,
                                LosColoradosRuntimePercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[7, 4 + i].Value.ToString())/100,
                                LosColoradosHoursHours = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[8, 4 + i].Value.ToString()),
                                LagunaSecaOreGradeCutPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[11, 4 + i].Value.ToString())/100,
                                LagunaSecaMillRecoveryPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[12, 4 + i].Value.ToString())/100,
                                LagunaSecaMillFeedTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[13, 4 + i].Value.ToString()) / 1000,
                                LagunaSecaCuExMillTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[14, 4 + i].Value.ToString()) / 1000,
                                LagunaSecaRuntimePercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[15, 4 + i].Value.ToString())/100,
                                LagunaSecaHoursHours = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[16, 4 + i].Value.ToString()),
                                LagunaSeca2OreGradeCutPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[19, 4 + i].Value.ToString())/100,
                                LagunaSeca2MillRecoveryPercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[20, 4 + i].Value.ToString())/100,
                                LagunaSeca2MillFeedTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[21, 4 + i].Value.ToString()) / 1000,
                                LagunaSeca2CuExMillTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[22, 4 + i].Value.ToString()) / 1000,
                                LagunaSeca2RuntimePercentage = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[23, 4 + i].Value.ToString())/100,
                                LagunaSeca2HoursHours = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[24, 4 + i].Value.ToString()),
                                OxideOreToOlTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[27, 4 + i].Value.ToString()) / 1000,
                                OxideCuCathodesTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[28, 4 + i].Value.ToString()) / 1000,
                                SulphideLeachStackedMaterialFromMineTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[29, 4 + i].Value.ToString()) / 1000,
                                SulphideLeachContractorsStackedMaterialFromStocksTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[30, 4 + i].Value.ToString()) / 1000,
                                SulphideLeachMelStackedMaterialFromStocksTonnesTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[31, 4 + i].Value.ToString()) / 1000,
                                SulphideLeachTotalStackedMaterialTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[32, 4 + i].Value.ToString()) / 1000,
                                SulphideLeachCuCathodesTonnes = double.Parse(budgetMovementProductionTemplateWorksheet.Cells[33, 4 + i].Value.ToString()) / 1000
                            });

                            _budgetPitDisintegrated.Add(new MineComplianceBudgetPitDisintegrated()
                            {
                                Date = _date,
                                ExpitEsTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 3].Value.ToString()) / 1000000,
                                ExpitEnTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 4].Value.ToString()) / 1000000,
                                TotalExpitTonnes = (double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 3].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 4].Value.ToString())) / 1000000,
                                MillRehandlingTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 5].Value.ToString()) / 1000000,
                                OlRehandlingTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 6].Value.ToString()) / 1000000,
                                SlRehandlingTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 7].Value.ToString()) / 1000000,
                                OtherRehandlingTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 8].Value.ToString()) / 1000000,
                                TotalRehandlingTonnes = (double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 5].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 6].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 7].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 8].Value.ToString())) / 1000000,
                                TotalMovementTonnes = (double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 3].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 4].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 5].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 6].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 7].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 8].Value.ToString())) / 1000000,
                                RehandlingTotalTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 9].Value.ToString()) / 1000000,
                                MovementTotalTonnes = double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 10].Value.ToString()) / 1000000,
                                TotalTonnes = (double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 9].Value.ToString()) + double.Parse(budgetPitDisintegratedTemplateWorksheet.Cells[6 + i, 10].Value.ToString())) / 1000000
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

                var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.MineComplianceExcelFilePath;
                var loadFileInfo = new FileInfo(loadFilePath);

                if (loadFileInfo.Exists)
                {
                    var package = new ExcelPackage(loadFileInfo);
                    var budgetPrincipalWorksheet = package.Workbook.Worksheets[MineComplianceConstants.BudgetPrincipalMineComplianceSpotfireWorksheet];
                    var budgetMovementProductionWorksheet = package.Workbook.Worksheets[MineComplianceConstants.BudgetMovementProductionMineComplianceSpotfireWorksheet];
                    var budgetPitDisintegratedWorksheet = package.Workbook.Worksheets[MineComplianceConstants.BudgetPitDisintegratedMineComplianceSpotfireWorksheet];

                    if (budgetPrincipalWorksheet != null & budgetMovementProductionWorksheet != null & budgetPitDisintegratedWorksheet != null)
                    {
                        try
                        {
                            var openWriteCheck = File.OpenWrite(loadFilePath);
                            openWriteCheck.Close();

                            var lastRow1 = budgetPrincipalWorksheet.Dimension.End.Row + 1;
                            for (var i = 0; i < _budgetPrincipal.Count; i++)
                            {
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 1].Value = _budgetPrincipal[i].Date;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 2].Value = _budgetPrincipal[i].ExpitTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 3].Value = _budgetPrincipal[i].RehandlingTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 4].Value = _budgetPrincipal[i].MovementTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 5].Value = _budgetPrincipal[i].ShovelsUnits73Yd3;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 6].Value = _budgetPrincipal[i].ShovelsAvailabilityPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 7].Value = _budgetPrincipal[i].ShovelsUtilizationPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 8].Value = _budgetPrincipal[i].ShovelsPerformanceTonnesPerHour;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 9].Value = _budgetPrincipal[i].ShovelsStandByHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 10].Value = _budgetPrincipal[i].ShovelsProductionTimeHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 11].Value = _budgetPrincipal[i].ShovelAvailableHoursHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 12].Value = _budgetPrincipal[i].ShovelHoursHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 13].Value = _budgetPrincipal[i].TrucksUnits;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 14].Value = _budgetPrincipal[i].TrucksAvailabilityPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 15].Value = _budgetPrincipal[i].TrucksUtilizationPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 16].Value = _budgetPrincipal[i].TrucksPerformanceTonnesPerDay;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 17].Value = _budgetPrincipal[i].TrucksStandByHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 18].Value = _budgetPrincipal[i].TrucksHoursHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 19].Value = _budgetPrincipal[i].TrucksProductionTimeHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 20].Value = _budgetPrincipal[i].TrucksAvailableHoursHours;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 21].Value = _budgetPrincipal[i].MillThroughputTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 22].Value = _budgetPrincipal[i].MillGradeCuPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 23].Value = _budgetPrincipal[i].MillRecoveryPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 24].Value = _budgetPrincipal[i].MillRehandlingPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 25].Value = _budgetPrincipal[i].OlThroughputTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 26].Value = _budgetPrincipal[i].OlGradeCuPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 27].Value = _budgetPrincipal[i].OlRecoveryPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 28].Value = _budgetPrincipal[i].OlCuSPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 29].Value = _budgetPrincipal[i].SlThroughputTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 30].Value = _budgetPrincipal[i].SlGradeCuPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 31].Value = _budgetPrincipal[i].SlRecoveryPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 32].Value = _budgetPrincipal[i].SlCuSPercentage;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 33].Value = _budgetPrincipal[i].MillProductionTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 34].Value = _budgetPrincipal[i].CathodesTonnes;
                                budgetPrincipalWorksheet.Cells[i + lastRow1, 35].Value = _budgetPrincipal[i].TotalProductionTonnes;
                            }

                            var lastRow2 = budgetMovementProductionWorksheet.Dimension.End.Row + 1;

                            for (var i = 0; i < _budgetMovementProduction.Count; i++)
                            {
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 1].Value = _budgetMovementProduction[i].Date;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 2].Value = _budgetMovementProduction[i].LosColoradosOreGradeCutPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 3].Value = _budgetMovementProduction[i].LosColoradosMillRecoveryPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 4].Value = _budgetMovementProduction[i].LosColoradosMillFeedTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 5].Value = _budgetMovementProduction[i].LosColoradosCuExMillTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 6].Value = _budgetMovementProduction[i].LosColoradosRuntimePercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 7].Value = _budgetMovementProduction[i].LosColoradosHoursHours;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 8].Value = _budgetMovementProduction[i].LagunaSecaOreGradeCutPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 9].Value = _budgetMovementProduction[i].LagunaSecaMillRecoveryPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 10].Value = _budgetMovementProduction[i].LagunaSecaMillFeedTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 11].Value = _budgetMovementProduction[i].LagunaSecaCuExMillTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 12].Value = _budgetMovementProduction[i].LagunaSecaRuntimePercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 13].Value = _budgetMovementProduction[i].LagunaSecaHoursHours;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 14].Value = _budgetMovementProduction[i].LagunaSeca2OreGradeCutPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 15].Value = _budgetMovementProduction[i].LagunaSeca2MillRecoveryPercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 16].Value = _budgetMovementProduction[i].LagunaSeca2MillFeedTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 17].Value = _budgetMovementProduction[i].LagunaSeca2CuExMillTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 18].Value = _budgetMovementProduction[i].LagunaSeca2RuntimePercentage;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 19].Value = _budgetMovementProduction[i].LagunaSeca2HoursHours;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 20].Value = _budgetMovementProduction[i].OxideOreToOlTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 21].Value = _budgetMovementProduction[i].OxideCuCathodesTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 22].Value = _budgetMovementProduction[i].SulphideLeachStackedMaterialFromMineTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 23].Value = _budgetMovementProduction[i].SulphideLeachContractorsStackedMaterialFromStocksTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 24].Value = _budgetMovementProduction[i].SulphideLeachMelStackedMaterialFromStocksTonnesTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 25].Value = _budgetMovementProduction[i].SulphideLeachTotalStackedMaterialTonnes;
                                budgetMovementProductionWorksheet.Cells[i + lastRow2, 26].Value = _budgetMovementProduction[i].SulphideLeachCuCathodesTonnes;
                            }

                            var lastRow3 = budgetPitDisintegratedWorksheet.Dimension.End.Row + 1;

                            for (var i = 0; i < _budgetPitDisintegrated.Count; i++)
                            {
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 1].Value = _budgetPitDisintegrated[i].Date;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 2].Value = _budgetPitDisintegrated[i].ExpitEsTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 3].Value = _budgetPitDisintegrated[i].ExpitEnTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 4].Value = _budgetPitDisintegrated[i].TotalExpitTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 5].Value = _budgetPitDisintegrated[i].MillRehandlingTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 6].Value = _budgetPitDisintegrated[i].OlRehandlingTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 7].Value = _budgetPitDisintegrated[i].SlRehandlingTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 8].Value = _budgetPitDisintegrated[i].OtherRehandlingTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 9].Value = _budgetPitDisintegrated[i].TotalRehandlingTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 10].Value = _budgetPitDisintegrated[i].TotalMovementTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 11].Value = _budgetPitDisintegrated[i].RehandlingTotalTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 12].Value = _budgetPitDisintegrated[i].MovementTotalTonnes;
                                budgetPitDisintegratedWorksheet.Cells[i + lastRow3, 13].Value = _budgetPitDisintegrated[i].TotalTonnes;
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