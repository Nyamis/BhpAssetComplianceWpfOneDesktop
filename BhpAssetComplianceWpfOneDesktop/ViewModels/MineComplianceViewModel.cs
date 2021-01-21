using BhpAssetComplianceWpfOneDesktop.Models;
using BhpAssetComplianceWpfOneDesktop.Resources;
using Prism.Mvvm;
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

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class MineComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.MineCompliance;

        private string _UpdateA;
        public string UpdateA
        {
            get { return _UpdateA; }
            set { SetProperty(ref _UpdateA, value); }
        }

        private string _UpdateB;
        public string UpdateB
        {
            get { return _UpdateB; }
            set { SetProperty(ref _UpdateB, value); }
        }

        DateTime _Date;
        public DateTime Date
        {
            get { return _Date; }
            set { SetProperty(ref _Date, value); }
        }

        private int _FiscalYear;
        public int FiscalYear
        {
            get { return _FiscalYear; }
            set { SetProperty(ref _FiscalYear, value); }
        }

        private bool _isEnabled1;
        public bool IsEnabled1
        {
            get { return _isEnabled1; }
            set { SetProperty(ref _isEnabled1, value); }
        }

        private bool _isEnabled2;
        public bool IsEnabled2
        {
            get { return _isEnabled2; }
            set { SetProperty(ref _isEnabled2, value); }
        }


        public DelegateCommand GenerarMCRT { get; private set; }
        public DelegateCommand CargarMCRT { get; private set; }
        public DelegateCommand GenerarMCBT { get; private set; }
        public DelegateCommand CargarMCBT { get; private set; }

        public MineComplianceViewModel()
        {
            Date = DateTime.Now;
            FiscalYear = Date.Year;
            IsEnabled1 = true;
            IsEnabled2 = false;
            GenerarMCRT = new DelegateCommand(GenerateRealTemplate);
            CargarMCRT = new DelegateCommand(LoadRealTemplate).ObservesCanExecute(() => IsEnabled1);
            GenerarMCBT = new DelegateCommand(GenerateBudgetTemplate);
            CargarMCBT = new DelegateCommand(LoadBudgetTemplate).ObservesCanExecute(() => IsEnabled2);
        }

        private void GenerateRealTemplate()
        {
            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Mine Compliance Real Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("Real Pit Disintegrated");
            List<string> lstCategory1 = new List<string>() { "Category", "Parameter", "Unit", "Month" };
            List<string> lstParameter1 = new List<string>() { "Expit ES", "Expit EN", "Mill", "OL", "SL", "Other", "Reh", "Mov." };

            ws.Cells["C2:J2"].Merge = true;
            ws.Cells["C2:J2"].Style.Font.Bold = true;
            ws.Cells["C2:J2"].Value = "Real";
            ws.Cells["C2:J2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells["C2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C2:J2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C2:J2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#002060"));

            ws.Cells[$"C2:J2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[$"B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[$"J2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < 5; i++)
            {
                ws.Cells[$"B{2 + i}:J{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] G = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
            for (int i = G.GetLowerBound(0); i <= G.GetUpperBound(0); i++)
            {
                ws.Cells[$"{G[i]}3:{G[i]}6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Column(1 + i).Width = 14;
            }

            //ws.Cells["A6"].Style.Font.Bold = true;
            //ws.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //ws.Cells["A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws.Cells["A6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));

            for (int i = 0; i < lstCategory1.Count; i++)
            {
                ws.Cells[3 + i, 2].Value = lstCategory1[i];
                ws.Cells[3 + i, 2].Style.Font.Bold = true;
                ws.Cells[3 + i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[3 + i, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            }

            ws.Cells["C3:D3"].Merge = true;
            ws.Cells["E3:H3"].Merge = true;
            ws.Cells["I3:J3"].Merge = true;
            ws.Cells["C3:J3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells[3, 3].Value = "Movement";
            ws.Cells[3, 5].Value = "Rehandling";
            ws.Cells[3, 9].Value = "Total";

            ws.Cells["C3:J3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
            ws.Cells["E3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E3:H3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

            for (int i = 0; i < 3; i++)
            {
                ws.Cells[$"C{3 + i}:J{3 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < lstParameter1.Count; i++)
            {
                ws.Cells[4, 3 + i].Value = lstParameter1[i];
                ws.Cells[5, 3 + i].Value = "t";
            }
            ws.Cells["C4:J4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws.Cells["C4:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C4:J4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
            ws.Cells["E4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E4:H4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

            ws.Cells["C5:J5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C5:J5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
            ws.Cells["E5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E5:H5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));

            //CultureInfo ci = new CultureInfo("en-US");
            //ws.Cells[6, 2].Value = $"{Date.ToString("MMMM", ci)}";

            var ws2 = pck.Workbook.Worksheets.Add("Real Movement Production");
            List<string> lstZone2 = new List<string>() { "Los Colorados", "Laguna Seca", "Laguna Seca 2", "Oxide", "Sulphide Leach", "Coloso" };
            List<string> lstParameter21 = new List<string>() { "Ore Grade - CuT", "Mill Recovery", "Mill Feed", "Cu Ex Mill", "Runtime", "Hours" };
            List<string> lstParameter22 = new List<string>() { "CuT Ore Grade (ROM)", "CuS Ore Grade (ROM)", "CuT Ore Grade (Crusher)", "CuS Ore Grade (Crusher)", "Recovery (Crusher + ROM)", "ROM", "Crushed Material", "Total Stacked Material", "Cu Cathodes", "CuT Ore Grade", "CuS Ore Grade", "Recovery", "Stacked Material from Mine", "Contractors Stacked Material from Stocks", "MEL Stacked Material from Stocks", "Total Stacked Material", "Cu Cathodes" };
            List<string> lstUnit21 = new List<string>() { "%", "%", "t", "t Cu", "%", "h" };
            List<string> lstUnit22 = new List<string>() { "%", "%", "%", "%", "%", "dmt", "dmt", "dmt", "t Cu", "%", "%", "%", "dmt", "dmt", "dmt", "dmt", "t Cu" };

            ws2.Column(1).Width = 19;
            int[] H = { 3, 11, 19, 27, 36, 46 };
            for (int i = H.GetLowerBound(0); i <= H.GetUpperBound(0); i++)
            {
                ws2.Cells[H[i], 1].Value = lstZone2[i];
                ws2.Cells[$"B{H[i]}:D{H[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws2.Cells[H[i], 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[H[i], 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                if (i % 2 != 0)
                {
                    ws2.Cells[H[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[H[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));
                }
                else if (i % 2 == 0)
                {
                    ws2.Cells[H[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[H[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
                }
            }

            ws2.Column(2).Width = 35;
            ws2.Column(2).Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            for (int i = 0; i < lstParameter21.Count; i++)
            {
                ws2.Cells[3 + i, 2].Value = lstParameter21[i];
                ws2.Cells[3 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{3 + i}:D{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ws2.Cells[11 + i, 2].Value = lstParameter21[i];
                ws2.Cells[11 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{11 + i}:D{11 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ws2.Cells[19 + i, 2].Value = lstParameter21[i];
                ws2.Cells[19 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{19 + i}:D{19 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            int[] I = { 3, 19, 36 };
            int[] J = { 8, 24, 43 };
            for (int i = I.GetLowerBound(0); i <= I.GetUpperBound(0); i++)
            {
                ws2.Cells[$"B{I[i]}:B{J[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"B{I[i]}:B{J[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                ws2.Cells[$"B{I[i]}:B{J[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"B{I[i]}:B{J[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws2.Cells[$"C{I[i]}:C{J[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"C{I[i]}:C{J[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws2.Cells[$"C{I[i]}:C{J[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"C{I[i]}:C{J[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"D{I[i]}:D{J[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            }

            int[] K = { 11, 27, 46 };
            int[] L = { 16, 35, 47 };
            for (int i = K.GetLowerBound(0); i <= K.GetUpperBound(0); i++)
            {
                ws2.Cells[$"B{K[i]}:B{L[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"B{K[i]}:B{L[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));
                ws2.Cells[$"B{K[i]}:B{L[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"B{K[i]}:B{L[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws2.Cells[$"C{K[i]}:C{L[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"C{K[i]}:C{L[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                ws2.Cells[$"C{K[i]}:C{L[i]}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"C{K[i]}:C{L[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"D{K[i]}:D{L[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            ws2.Column(3).Width = 11;
            ws2.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws2.Column(4).Width = 11;

            int[] M = { 2, 10, 18, 26, 45 };
            for (int i = 0; i < 5; i++)
            {
                ws2.Cells[M[i], 3].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws2.Cells[M[i], 3].Value = "Unit";
                ws2.Cells[M[i], 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws2.Cells[M[i], 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[M[i], 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                if (i % 2 != 0)
                {
                    ws2.Cells[M[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[M[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));
                }
                else if (i % 2 == 0)
                {
                    ws2.Cells[M[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[M[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                }
            }
            ws2.Cells[45, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells[45, 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

            for (int i = 0; i < lstUnit22.Count; i++)
            {
                ws2.Cells[27 + i, 3].Value = lstUnit22[i];
                ws2.Cells[27 + i, 2].Value = lstParameter22[i];

                ws2.Cells[$"B{27 + i}:D{27 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            ws2.Cells[46, 2].Value = "Cu Ex Coloso";
            ws2.Cells[46, 3].Value = "t";
            ws2.Cells["B46:D46"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws2.Cells[47, 2].Value = "Cu from low grade Concentrate";
            ws2.Cells[47, 3].Value = "t";
            ws2.Cells["B47:D47"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //ws2.Cells[1, 4].Value = "Month";
            //ws2.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws2.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));

            ws2.Cells[2, 4].Value = "Month";
            ws2.Cells[2, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells[2, 4].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            ws2.Cells["D2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws2.Cells["D2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws2.Cells["D2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws2.Cells["D2"].Style.Font.Bold = true;
            ws2.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var ws3 = pck.Workbook.Worksheets.Add("Loading FC");
            List<string> lstParameter4 = new List<string>() { "Units", "Availability", "Utilization (New TUM)", "Total Hours", "Available Hours", "Equipment scheduled downtime", "Equipment non-scheduled downtime", "Process scheduled downtime", "Process non-scheduled downtime", "Stand By", "Hang Time", "Production time(New TUM)", "Performance (New TUM)", "Total Tonnes" };
            List<string> lstHeader4 = new List<string>() { "SHOVEL BUCYRUS 495B", "SHOVEL P & H 4100XPB", "SHOVEL P & H 4100XPC", "SHOVEL BUCYRUS 495HR", "FRONT LOADER CAT 994F", "PC5500", "PC8000", "SUMMARY SHOVEL 73 yd3", "TOTAL FLEET" };
            List<string> lstUnit4 = new List<string>() { "N°", "%", "%", "h", "h", "h", "h", "h", "h", "h", "h", "h", "t/h", "kt" };

            ws3.Column(1).Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws3.Column(1).Width = 36;
            ws3.Cells["A1:A2"].Merge = true;
            ws3.Cells["A1"].Value = "Total Loading Fleet";
            ws3.Cells["A1:C1"].Style.Font.Bold = true;
            ws3.Cells["A1"].Style.Font.Size = 16;
            ws3.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#000000"));
            ws3.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws3.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Column(2).Width = 16;
            //ws3.Cells["B1:C1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws3.Cells["B2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws3.Cells["B1:B2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws3.Cells["B1:B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //ws3.Cells["C1:C2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws3.Cells["C4"].Style.Font.Bold = true;
            ws3.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws3.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws3.Cells["C4"].Value = "Month";
            //ws3.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws3.Cells["B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
            //ws3.Cells["B2"].Value = "Number of days";
            //ws3.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws3.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            //ws3.Cells[1, 3].Value = $"{Date.ToString("MMM", ci)}-{Date.ToString("yy")}";
            //ws3.Cells[2, 3].Value = $"{DateTime.DaysInMonth(Convert.ToInt32(Date.Year), Convert.ToInt32(Date.Month))}";
            //ws3.Cells["C1:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws3.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["C4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws3.Cells["C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws3.Cells["C4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


            int[] N = { 4, 20, 36, 52, 68, 84, 100, 116, 132 };
            for (int i = N.GetLowerBound(0); i <= N.GetUpperBound(0); i++)
            {
                ws3.Cells[N[i], 1].Value = lstHeader4[i];
                ws3.Cells[N[i], 1].Style.Font.Bold = true;
                ws3.Cells[N[i], 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells[N[i], 2].Value = "Unit";
                ws3.Cells[N[i], 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                ws3.Cells[$"A{N[i]}:B{N[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws3.Cells[$"A{N[i]}:c{N[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws3.Cells[$"A{N[i]}:B{N[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                if (i % 2 != 0)
                {
                    ws3.Cells[N[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

                    ws3.Cells[N[i] + 14, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i] + 14, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

                    ws3.Cells[N[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                    ws3.Cells[$"A{N[i] + 1 }:A{N[i] + 13 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[$"A{N[i] + 1 }:A{N[i] + 13 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                }
                else if (i % 2 == 0)
                {
                    ws3.Cells[N[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));

                    ws3.Cells[N[i] + 14, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i] + 14, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));

                    ws3.Cells[N[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[N[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                    ws3.Cells[$"A{N[i] + 1 }:A{N[i] + 13 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws3.Cells[$"A{N[i] + 1 }:A{N[i] + 13 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                }

                for (int j = 0; j < lstParameter4.Count; j++)
                {
                    ws3.Cells[N[i] + 1 + j, 1].Value = lstParameter4[j];
                    ws3.Cells[N[i] + 1 + j, 2].Value = lstUnit4[j];
                    ws3.Cells[$"A{N[i] + 1 + j}:C{N[i] + 1 + j}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws3.Cells[$"A{N[i] + 1 + j}:C{N[i] + 1 + j}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    if (i % 2 != 0)
                    {
                        ws3.Cells[N[i] + 1 + j, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws3.Cells[N[i] + 1 + j, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                    }
                    else if (i % 2 == 0)
                    {
                        ws3.Cells[N[i] + 1 + j, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws3.Cells[N[i] + 1 + j, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                    }

                }
            }
            ws3.Column(3).Width = 11;

            var ws4 = pck.Workbook.Worksheets.Add("Hauling FC");
            List<string> lstParameter5 = new List<string>() { "Units", "Mechanical Availability", "Physical Availability", "Utilization (New TUM)", "Total Hours", "Available Hours", "Equipment scheduled downtime", "Equipment non-scheduled downtime", "Process scheduled downtime", "Process non-scheduled downtime", "Standby", "Queue Time", "Production time (New TUM)", "Performance (New TUM)", "Cycle Time", "Total Tonnes" };
            List<string> lstHeader5 = new List<string>() { "930 Fleet", "960 Autonomous Fleet", "960 MEL Fleet", "Liebherr ESTRS", "Komatsu ESTRS", "CAT ESTRS", "797B MARC Fleet", "797B MEL Fleet", "797F MARC Fleet", "793F MEL Fleet", "240 Total Fleet (793 + 930)", "350 Total Fleet (797 + 960)", "Total Fleet" };
            List<string> lstUnit5 = new List<string>() { "N°", "%", "%", "%", "h", "h", "h", "h", "h", "h", "h", "h", "h", "t/h", "min", "kt" };

            ws4.Column(1).Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws4.Column(1).Width = 31;
            ws4.Cells["A1:A2"].Merge = true;
            ws4.Cells["A1"].Value = "Total Hauling Fleet";
            ws4.Cells["A1:C1"].Style.Font.Bold = true;
            ws4.Cells["A1"].Style.Font.Size = 16;
            ws4.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#000000"));
            ws4.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws4.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws4.Column(2).Width = 16;
            //ws4.Cells["B1:C1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws4.Cells["B2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws4.Cells["B1:B2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws4.Cells["B1:B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //ws4.Cells["C1:C2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws4.Cells["C4"].Style.Font.Bold = true;
            ws4.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws4.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws4.Cells["C4"].Value = "Month";
            //ws4.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws4.Cells["B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
            //ws4.Cells["B2"].Value = "Number of days";
            //ws4.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws4.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            //ws4.Cells[1, 3].Value = $"{Date.ToString("MMM", ci)}-{Date.ToString("yy")}";
            //ws4.Cells[2, 3].Value = $"{DateTime.DaysInMonth(Convert.ToInt32(Date.Year), Convert.ToInt32(Date.Month))}";
            //ws4.Cells["C1:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            ws4.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws4.Cells["C4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws4.Cells["C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws4.Cells["C4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            int[] O = { 4, 22, 40, 58, 76, 94, 112, 130, 148, 166, 184, 202, 220 };
            for (int i = O.GetLowerBound(0); i <= O.GetUpperBound(0); i++)
            {
                ws4.Cells[O[i], 1].Value = lstHeader5[i];
                ws4.Cells[O[i], 1].Style.Font.Bold = true;
                ws4.Cells[O[i], 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws4.Cells[O[i], 2].Value = "Unit";
                ws4.Cells[O[i], 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws4.Cells[O[i] + 16, 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                ws4.Cells[$"A{O[i]}:B{O[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws4.Cells[$"A{O[i]}:C{O[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws4.Cells[$"A{O[i]}:B{O[i]}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                if (i % 2 != 0)
                {
                    ws4.Cells[O[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

                    ws4.Cells[O[i] + 16, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i] + 16, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

                    ws4.Cells[O[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                    ws4.Cells[O[i] + 16, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i] + 16, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                    ws4.Cells[$"A{O[i] + 1 }:A{O[i] + 15 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[$"A{O[i] + 1 }:A{O[i] + 15 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));


                    ws4.Cells[$"B{O[i] + 1 }:B{O[i] + 15}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[$"B{O[i] + 1 }:B{O[i] + 15}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                }
                else if (i % 2 == 0)
                {
                    ws4.Cells[O[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));

                    ws4.Cells[O[i] + 16, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i] + 16, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));

                    ws4.Cells[O[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                    ws4.Cells[O[i] + 16, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[O[i] + 16, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                    ws4.Cells[$"A{O[i] + 1 }:A{O[i] + 15 }"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[$"A{O[i] + 1 }:A{O[i] + 15 }"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                    ws4.Cells[$"B{O[i] + 1 }:B{O[i] + 15}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws4.Cells[$"B{O[i] + 1 }:B{O[i] + 15}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                }

                for (int j = 0; j < lstParameter5.Count; j++)
                {
                    ws4.Cells[O[i] + 1 + j, 1].Value = lstParameter5[j];
                    ws4.Cells[O[i] + 1 + j, 2].Value = lstUnit5[j];
                    ws4.Cells[$"A{O[i] + 1 + j}:C{O[i] + 1 + j}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws4.Cells[$"A{O[i] + 1 + j}:C{O[i] + 1 + j}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
            }
            ws4.Column(3).Width = 11;

            var ws5 = pck.Workbook.Worksheets.Add("Mill FC");
            List<string> lstParameter3 = new List<string>() { "Ore Grade - CuT", "Mill Recovery", "Mill Feed" };
            List<string> lstUnit3 = new List<string>() { "%", "%", "dmt" };

            ws5.Cells["A1:A2"].Merge = true;
            ws5.Cells["A1"].Value = "Mill";
            ws5.Cells["A1"].Style.Font.Bold = true;
            ws5.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws5.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws5.Cells["A1"].Style.Font.Size = 16;

            //ws5.Cells["B1:C1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws5.Cells["B2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws5.Cells["B1:B2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws5.Cells["B1:B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //ws5.Cells["C1:C2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws5.Cells["B3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws5.Cells["B3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws5.Cells["B3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

            ws5.Cells["C3"].Style.Font.Bold = true;
            ws5.Cells["B3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws5.Cells["B3"].Value = "Item";
            ws5.Cells["C3"].Value = "Month";

            //ws5.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws5.Cells["B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));

            //ws5.Cells["B2"].Value = "Number of days";
            //ws5.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws5.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            //ws5.Cells[1, 3].Value = $"{Date.ToString("MMM", ci)}-{Date.ToString("yy")}";
            //ws5.Cells[2, 3].Value = $"{DateTime.DaysInMonth(Convert.ToInt32(Date.Year), Convert.ToInt32(Date.Month))}";
            //ws5.Cells["C1:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws5.Cells[$"A4:C4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            ws5.Cells["C3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws5.Cells["C3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws5.Cells["B3:C3"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws5.Cells["A3:C3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < lstParameter3.Count; i++)
            {
                ws5.Column(1 + i).Width = 16;
                ws5.Cells[4 + i, 1].Value = lstParameter3[i];
                ws5.Cells[4 + i, 2].Value = lstUnit3[i];
                ws5.Cells[$"A{4 + i}:C{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws5.Cells[$"A{4 + i}:C{4 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            ws5.Cells["A4:A6"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws5.Cells["A4:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws5.Cells["A4:A6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
            ws5.Cells["B4:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws5.Cells["B4:B6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));

            //if (Date.Month == 1 || Date.Month == 2 || Date.Month == 3 || Date.Month == 4 || Date.Month == 5 || Date.Month == 6)
            //{
            //    ws.Cells["A4"].Value = $"{Date.ToString("MMM")} FY{Date.Year}".ToUpper();
            //}
            //else if (Date.Month == 9)
            //{
            //    ws.Cells["A4"].Value = $"SEP. FY{Date.Year + 1}".ToUpper();
            //}
            //else
            //{
            //    ws.Cells["A4"].Value = $"{Date.ToString("MMM")} FY{Date.Year + 1}".ToUpper();
            //}

            //ws.Cells[2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            //ws.Cells[2, i + 1].Style.WrapText = true;

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "MineComplianceRealTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                FileStream fs = File.OpenWrite(dialog.FileName);
                fs.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabled1 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Upload Error");
            }
        }

        public class RealMovementProduction
        {
            public double LosColoradosOreGradeCutPercentage { get; set; }
            public double LosColoradosMillRecoveryPercentage { get; set; }
            public double LosColoradosMillFeedTonnes { get; set; }
            public double LosColoradosCuExMillTonnes { get; set; }
            public double LosColoradosRuntimePercentage { get; set; }
            public double LosColoradosHoursHours { get; set; }

            public double LagunaSecaOreGradeCutPercentage { get; set; }
            public double LagunaSecaMillRecoveryPercentage { get; set; }
            public double LagunaSecaMillFeedTonnes { get; set; }
            public double LagunaSecaCuExMillTonnes { get; set; }
            public double LagunaSecaRuntimePercentage { get; set; }
            public double LagunaSecaHoursHours { get; set; }

            public double LagunaSeca2OreGradeCutPercentage { get; set; }
            public double LagunaSeca2MillRecoveryPercentage { get; set; }
            public double LagunaSeca2MillFeedTonnes { get; set; }
            public double LagunaSeca2CuExMillTonnes { get; set; }
            public double LagunaSeca2RuntimePercentage { get; set; }
            public double LagunaSeca2HoursHours { get; set; }

            public double OxideCutOreGradeRomPercentage { get; set; }
            public double OxideCusOreGradeRomPercentage { get; set; }
            public double OxideCutOreGradeCrusherPercentage { get; set; }
            public double OxideCusOreGradeCrusherPercentage { get; set; }
            public double OxideRecoveryCrusherAndRomPercentage { get; set; }
            public double OxideRomTonnes { get; set; }
            public double OxideCrushedMaterialTonnes { get; set; }
            public double OxideTotalStackedMaterialTonnes { get; set; }
            public double OxideCuCathodesTonnes { get; set; }

            public double SulphideLeachCutOreGradePercentage { get; set; }
            public double SulphideLeachCusOreGradePercentage { get; set; }
            public double SulphideLeachRecoveryPercentage { get; set; }
            public double SulphideLeachStackedMaterialFromMineTonnes { get; set; }
            public double SulphideLeachContractorsStackedMaterialFromStocksTonnes { get; set; }
            public double SulphideLeachMelStackedMaterialFromStocksTonnes { get; set; }
            public double SulphideLeachTotalStackedMaterialTonnes { get; set; }
            public double SulphideLeachCuCathodesTonnes { get; set; }

            public double ColosoCuExColosoTonnes { get; set; }
            public double ColosoCuFromLowGradeConcentrateTonnes { get; set; }         
        }

        readonly List<RealMovementProduction> lstRealMovementProduction = new List<RealMovementProduction>();

        public class RealPitDisintegrated
        {
            public double ExpitEsTonnes { get; set; }
            public double ExpitEnTonnes { get; set; }
            public double TotalExpitTonnes { get; set; }

            public double MillRehandlingTonnes { get; set; }
            public double OlRehandlingTonnes { get; set; }
            public double SlRehandlingTonnes { get; set; }
            public double OtherRehandlingTonnes { get; set; }
            public double TotalRehandlingTonnes { get; set; }

            public double TotalMovementTonnes { get; set; }

            public double RehandlingTotalTonnes { get; set; }
            public double MovementTotalTonnes { get; set; }

            public double TotalTonnes { get; set; }

            //public double MillProductionTonnes { get; set; }
            //public double CathodesProductionTonnes { get; set; }

            //public double TotalProductionTonnes { get; set; }
        }
        
        readonly List<RealPitDisintegrated> lstRealPitDisintegrated = new List<RealPitDisintegrated>();

        public class RealMillFc
        {
            public double? OreGradeCut { get; set; }
            public double? MillRecovery { get; set; }
            public double? MillFeed { get; set; }
        }

        readonly List<RealMillFc> lstRealMillFc = new List<RealMillFc>();

        public class RealLoadingFc
        {          
            public string Name { get; set; }
            public double Units { get; set; }
            public double AvailabilityPercentage { get; set; }
            public double UtilizationPercentage { get; set; }
            public double TotalHoursHours { get; set; }
            public double AvailableHoursHours { get; set; }
            public double EquipmentScheduledDowntimeHours { get; set; }
            public double EquipmentNonScheduledDowntimeHours { get; set; }
            public double ProcessScheduledDowntimeHours { get; set; }
            public double ProcessNonScheduledDowntimeHours { get; set; }
            public double StandByHours { get; set; }
            public double HangTimeHours { get; set; }
            public double ProductionTimeHours { get; set; }
            public double PerformanceTonnesPerHour { get; set; }
            public double TotalTonnesTonnes { get; set; }
        }

        readonly List<RealLoadingFc> lstRealLoadingFc = new List<RealLoadingFc>();

        public class RealHaulingFc
        {
            public string Name { get; set; }
            public double Units { get; set; }
            public double MechanicalAvailabilityPercentage { get; set; }
            public double PhysicalAvailabilityPercentage { get; set; }
            public double UtilizationPercentage { get; set; }
            public double TotalHoursHours { get; set; }
            public double AvailableHoursHours { get; set; }
            public double EquipmentScheduledDowntimeHours { get; set; }
            public double EquipmentNonScheduledDowntimeHours { get; set; }
            public double ProcessScheduledDowntimeHours { get; set; }
            public double ProcessNonScheduledDowntimeHours { get; set; }
            public double StandByHours { get; set; }
            public double QueueTimeHours { get; set; }
            public double ProductionTimeHoursHours { get; set; }
            public double PerformanceTonnesPerHour { get; set; }
            public double CycleTimeHours { get; set; }
            public double TotalTonnesTonnes { get; set; }
        }

        readonly List<RealHaulingFc> lstRealHaulingFc = new List<RealHaulingFc>();

        private void LoadRealTemplate()
        {
            lstRealMovementProduction.Clear();
            lstRealPitDisintegrated.Clear();
            lstRealMillFc.Clear();
            lstRealLoadingFc.Clear();
            lstRealHaulingFc.Clear();

            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (op.ShowDialog() == true)
            {

                try
                {
                    FileInfo FilePath = new FileInfo(op.FileName);
                    ExcelPackage pck = new ExcelPackage(FilePath);

                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    ExcelWorksheet ws = pck.Workbook.Worksheets["Real Movement Production"];

                    for (int i = 0; i < 45; i++)
                    {
                        if (ws.Cells[3 + i, 4].Value == null)
                        {
                            ws.Cells[3 + i, 4].Value = 0;
                        }
                    }

                    lstRealMovementProduction.Add(new RealMovementProduction()
                    {
                        LosColoradosOreGradeCutPercentage = double.Parse(ws.Cells[3, 4].Value.ToString()),
                        LosColoradosMillRecoveryPercentage = double.Parse(ws.Cells[4, 4].Value.ToString()),
                        LosColoradosMillFeedTonnes = double.Parse(ws.Cells[5, 4].Value.ToString()) / 1000,
                        LosColoradosCuExMillTonnes = double.Parse(ws.Cells[6, 4].Value.ToString()) / 1000,
                        LosColoradosRuntimePercentage = double.Parse(ws.Cells[7, 4].Value.ToString()),
                        LosColoradosHoursHours = double.Parse(ws.Cells[8, 4].Value.ToString()),
                        LagunaSecaOreGradeCutPercentage = double.Parse(ws.Cells[11, 4].Value.ToString()),
                        LagunaSecaMillRecoveryPercentage = double.Parse(ws.Cells[12, 4].Value.ToString()),
                        LagunaSecaMillFeedTonnes = double.Parse(ws.Cells[13, 4].Value.ToString()) / 1000,
                        LagunaSecaCuExMillTonnes = double.Parse(ws.Cells[14, 4].Value.ToString()) / 1000,
                        LagunaSecaRuntimePercentage = double.Parse(ws.Cells[15, 4].Value.ToString()),
                        LagunaSecaHoursHours = double.Parse(ws.Cells[16, 4].Value.ToString()),
                        LagunaSeca2OreGradeCutPercentage = double.Parse(ws.Cells[19, 4].Value.ToString()),
                        LagunaSeca2MillRecoveryPercentage = double.Parse(ws.Cells[20, 4].Value.ToString()),
                        LagunaSeca2MillFeedTonnes = double.Parse(ws.Cells[21, 4].Value.ToString()) / 1000,
                        LagunaSeca2CuExMillTonnes = double.Parse(ws.Cells[22, 4].Value.ToString()) / 1000,
                        LagunaSeca2RuntimePercentage = double.Parse(ws.Cells[23, 4].Value.ToString()),
                        LagunaSeca2HoursHours = double.Parse(ws.Cells[24, 4].Value.ToString()),
                        OxideCutOreGradeRomPercentage = double.Parse(ws.Cells[27, 4].Value.ToString()),
                        OxideCusOreGradeRomPercentage = double.Parse(ws.Cells[28, 4].Value.ToString()),
                        OxideCutOreGradeCrusherPercentage = double.Parse(ws.Cells[29, 4].Value.ToString()),
                        OxideCusOreGradeCrusherPercentage = double.Parse(ws.Cells[30, 4].Value.ToString()),
                        OxideRecoveryCrusherAndRomPercentage = double.Parse(ws.Cells[31, 4].Value.ToString()),
                        OxideRomTonnes = double.Parse(ws.Cells[32, 4].Value.ToString()) / 1000000,
                        OxideCrushedMaterialTonnes = double.Parse(ws.Cells[33, 4].Value.ToString()) / 1000,
                        OxideTotalStackedMaterialTonnes = double.Parse(ws.Cells[34, 4].Value.ToString()) / 1000000,
                        OxideCuCathodesTonnes = double.Parse(ws.Cells[35, 4].Value.ToString()) / 1000,
                        SulphideLeachCutOreGradePercentage = double.Parse(ws.Cells[36, 4].Value.ToString()),
                        SulphideLeachCusOreGradePercentage = double.Parse(ws.Cells[37, 4].Value.ToString()),
                        SulphideLeachRecoveryPercentage = double.Parse(ws.Cells[38, 4].Value.ToString()),
                        SulphideLeachStackedMaterialFromMineTonnes = double.Parse(ws.Cells[39, 4].Value.ToString()) / 1000,
                        SulphideLeachContractorsStackedMaterialFromStocksTonnes = double.Parse(ws.Cells[40, 4].Value.ToString()) / 1000,
                        SulphideLeachMelStackedMaterialFromStocksTonnes = double.Parse(ws.Cells[41, 4].Value.ToString()) / 1000,
                        SulphideLeachTotalStackedMaterialTonnes = double.Parse(ws.Cells[42, 4].Value.ToString()) / 1000000,
                        SulphideLeachCuCathodesTonnes = double.Parse(ws.Cells[43, 4].Value.ToString()) / 1000,
                        ColosoCuExColosoTonnes = double.Parse(ws.Cells[46, 4].Value.ToString()) / 1000,
                        ColosoCuFromLowGradeConcentrateTonnes = double.Parse(ws.Cells[47, 4].Value.ToString()) / 1000
                    });

                    ExcelWorksheet ws2 = pck.Workbook.Worksheets["Real Pit Disintegrated"];

                    lstRealPitDisintegrated.Add(new RealPitDisintegrated()
                    {
                        ExpitEsTonnes = double.Parse(ws2.Cells[6, 3].Value.ToString()) / 1000000,
                        ExpitEnTonnes = double.Parse(ws2.Cells[6, 4].Value.ToString()) / 1000000,
                        TotalExpitTonnes = (double.Parse(ws2.Cells[6, 3].Value.ToString()) + double.Parse(ws2.Cells[6, 4].Value.ToString())) / 1000000,

                        MillRehandlingTonnes = double.Parse(ws2.Cells[6, 5].Value.ToString()) / 1000000,
                        OlRehandlingTonnes = double.Parse(ws2.Cells[6, 6].Value.ToString()) / 1000000,
                        SlRehandlingTonnes = double.Parse(ws2.Cells[6, 7].Value.ToString()) / 1000000,
                        OtherRehandlingTonnes = double.Parse(ws2.Cells[6, 8].Value.ToString()) / 1000000,
                        TotalRehandlingTonnes = (double.Parse(ws2.Cells[6, 5].Value.ToString()) + double.Parse(ws2.Cells[6, 6].Value.ToString()) + double.Parse(ws2.Cells[6, 7].Value.ToString()) + double.Parse(ws2.Cells[6, 8].Value.ToString())) / 1000000,

                        TotalMovementTonnes = (double.Parse(ws2.Cells[6, 3].Value.ToString()) + double.Parse(ws2.Cells[6, 4].Value.ToString()) + double.Parse(ws2.Cells[6, 5].Value.ToString()) + double.Parse(ws2.Cells[6, 6].Value.ToString()) + double.Parse(ws2.Cells[6, 7].Value.ToString()) + double.Parse(ws2.Cells[6, 8].Value.ToString())) / 1000000,

                        RehandlingTotalTonnes = double.Parse(ws2.Cells[6, 9].Value.ToString()) / 1000000,
                        MovementTotalTonnes = double.Parse(ws2.Cells[6, 10].Value.ToString()) / 1000000,
                        TotalTonnes = (double.Parse(ws2.Cells[6, 9].Value.ToString()) + double.Parse(ws2.Cells[6, 10].Value.ToString())) / 1000000
                    });

                    ExcelWorksheet ws3 = pck.Workbook.Worksheets["Mill FC"];

                    lstRealMillFc.Add(new RealMillFc()
                    {
                        OreGradeCut = double.Parse(ws3.Cells[4, 3].Value.ToString()),
                        MillRecovery = double.Parse(ws3.Cells[5, 3].Value.ToString()),
                        MillFeed = double.Parse(ws3.Cells[6, 3].Value.ToString()) / 1000000
                    });

                    ExcelWorksheet ws4 = pck.Workbook.Worksheets["Loading FC"];

                    //int rows = ws4.Dimension.Rows;
                    int[] a = { 0, 16, 32, 48, 64, 80, 96, 112, 128 };
                    for (int i = a.GetLowerBound(0); i <= a.GetUpperBound(0); i++)
                    {

                        for (int j = 1; j < 15; j++)
                        {

                            if (ws4.Cells[4 + a[i] + j, 3].Value == null)
                            {
                                ws4.Cells[4 + a[i] + j, 3].Value = 0;
                            }
                        }


                        lstRealLoadingFc.Add(new RealLoadingFc()
                        {
                            Name = ws4.Cells[4 + a[i], 1].Value.ToString(),
                            Units = double.Parse(ws4.Cells[5 + a[i], 3].Value.ToString()),
                            AvailabilityPercentage = double.Parse(ws4.Cells[6 + a[i], 3].Value.ToString()),
                            UtilizationPercentage = double.Parse(ws4.Cells[7 + a[i], 3].Value.ToString()),
                            TotalHoursHours = double.Parse(ws4.Cells[8 + a[i], 3].Value.ToString()),
                            AvailableHoursHours = double.Parse(ws4.Cells[9 + a[i], 3].Value.ToString()),
                            EquipmentScheduledDowntimeHours = double.Parse(ws4.Cells[10 + a[i], 3].Value.ToString()),
                            EquipmentNonScheduledDowntimeHours = double.Parse(ws4.Cells[11 + a[i], 3].Value.ToString()),
                            ProcessScheduledDowntimeHours = double.Parse(ws4.Cells[12 + a[i], 3].Value.ToString()),
                            ProcessNonScheduledDowntimeHours = double.Parse(ws4.Cells[13 + a[i], 3].Value.ToString()),
                            StandByHours = double.Parse(ws4.Cells[14 + a[i], 3].Value.ToString()),
                            HangTimeHours = double.Parse(ws4.Cells[15 + a[i], 3].Value.ToString()),
                            ProductionTimeHours = double.Parse(ws4.Cells[16 + a[i], 3].Value.ToString()),

                            PerformanceTonnesPerHour = double.Parse(ws4.Cells[17 + a[i], 3].Value.ToString()) / 1000000,
                            TotalTonnesTonnes = double.Parse(ws4.Cells[18 + a[i], 3].Value.ToString()) / 1000
                        });
                    }

                    ExcelWorksheet ws5 = pck.Workbook.Worksheets["Hauling FC"];

                    int[] b = { 0, 18, 36, 54, 72, 90, 108, 126, 144, 162, 180, 198, 216 };

                    for (int i = b.GetLowerBound(0); i <= b.GetUpperBound(0); i++)
                    {

                        for (int j = 1; j < 21; j++)
                        {
                            if (ws5.Cells[4 + b[i] + j, 3].Value == null)
                            {
                                ws5.Cells[4 + b[i] + j, 3].Value = 0;
                            }
                        }

                        lstRealHaulingFc.Add(new RealHaulingFc()
                        {
                            Name = ws5.Cells[4 + b[i], 1].Value.ToString(),
                            Units = double.Parse(ws5.Cells[5 + b[i], 3].Value.ToString()),
                            MechanicalAvailabilityPercentage = double.Parse(ws5.Cells[6 + b[i], 3].Value.ToString()),
                            PhysicalAvailabilityPercentage = double.Parse(ws5.Cells[7 + b[i], 3].Value.ToString()),
                            UtilizationPercentage = double.Parse(ws5.Cells[8 + b[i], 3].Value.ToString()),
                            TotalHoursHours = double.Parse(ws5.Cells[9 + b[i], 3].Value.ToString()),
                            AvailableHoursHours = double.Parse(ws5.Cells[10 + b[i], 3].Value.ToString()),
                            EquipmentScheduledDowntimeHours = double.Parse(ws5.Cells[11 + b[i], 3].Value.ToString()),
                            EquipmentNonScheduledDowntimeHours = double.Parse(ws5.Cells[12 + b[i], 3].Value.ToString()),
                            ProcessScheduledDowntimeHours = double.Parse(ws5.Cells[13 + b[i], 3].Value.ToString()),
                            ProcessNonScheduledDowntimeHours = double.Parse(ws5.Cells[14 + b[i], 3].Value.ToString()),
                            StandByHours = double.Parse(ws5.Cells[15 + b[i], 3].Value.ToString()),
                            QueueTimeHours = double.Parse(ws5.Cells[16 + b[i], 3].Value.ToString()),
                            ProductionTimeHoursHours = double.Parse(ws5.Cells[17 + b[i], 3].Value.ToString()),
                            PerformanceTonnesPerHour = double.Parse(ws5.Cells[18 + b[i], 3].Value.ToString()) / 1000000,
                            CycleTimeHours = double.Parse(ws5.Cells[19 + b[i], 3].Value.ToString()) / 60,
                            TotalTonnesTonnes = double.Parse(ws5.Cells[20 + b[i], 3].Value.ToString()) / 1000,
                        });
                    }

                    pck.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\MineComplianceData.xlsx";
                FileInfo filePath = new FileInfo(fileName);
                
                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00).AddMilliseconds(000);

                        ExcelWorksheet ws21 = pck2.Workbook.Worksheets["RealMovementProduction"];
                        int lastRow1 = ws21.Dimension.End.Row + 1;
                        ws21.Cells[lastRow1, 1].Value = newDate;
                        ws21.Cells[lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws21.Cells[lastRow1, 2].Value = lstRealMovementProduction[0].LosColoradosOreGradeCutPercentage;
                        ws21.Cells[lastRow1, 3].Value = lstRealMovementProduction[0].LosColoradosMillRecoveryPercentage;
                        ws21.Cells[lastRow1, 4].Value = lstRealMovementProduction[0].LosColoradosMillFeedTonnes;
                        ws21.Cells[lastRow1, 5].Value = lstRealMovementProduction[0].LosColoradosCuExMillTonnes;
                        ws21.Cells[lastRow1, 6].Value = lstRealMovementProduction[0].LosColoradosRuntimePercentage;
                        ws21.Cells[lastRow1, 7].Value = lstRealMovementProduction[0].LosColoradosHoursHours;
                        ws21.Cells[lastRow1, 8].Value = lstRealMovementProduction[0].LagunaSecaOreGradeCutPercentage;
                        ws21.Cells[lastRow1, 9].Value = lstRealMovementProduction[0].LagunaSecaMillRecoveryPercentage;
                        ws21.Cells[lastRow1, 10].Value = lstRealMovementProduction[0].LagunaSecaMillFeedTonnes;
                        ws21.Cells[lastRow1, 11].Value = lstRealMovementProduction[0].LagunaSecaCuExMillTonnes;
                        ws21.Cells[lastRow1, 12].Value = lstRealMovementProduction[0].LagunaSecaRuntimePercentage;
                        ws21.Cells[lastRow1, 13].Value = lstRealMovementProduction[0].LagunaSecaHoursHours;
                        ws21.Cells[lastRow1, 14].Value = lstRealMovementProduction[0].LagunaSeca2OreGradeCutPercentage;
                        ws21.Cells[lastRow1, 15].Value = lstRealMovementProduction[0].LagunaSeca2MillRecoveryPercentage;
                        ws21.Cells[lastRow1, 16].Value = lstRealMovementProduction[0].LagunaSeca2MillFeedTonnes;
                        ws21.Cells[lastRow1, 17].Value = lstRealMovementProduction[0].LagunaSeca2CuExMillTonnes;
                        ws21.Cells[lastRow1, 18].Value = lstRealMovementProduction[0].LagunaSeca2RuntimePercentage;
                        ws21.Cells[lastRow1, 19].Value = lstRealMovementProduction[0].LagunaSeca2HoursHours;
                        ws21.Cells[lastRow1, 20].Value = lstRealMovementProduction[0].OxideCutOreGradeRomPercentage;
                        ws21.Cells[lastRow1, 21].Value = lstRealMovementProduction[0].OxideCusOreGradeRomPercentage;
                        ws21.Cells[lastRow1, 22].Value = lstRealMovementProduction[0].OxideCutOreGradeCrusherPercentage;
                        ws21.Cells[lastRow1, 23].Value = lstRealMovementProduction[0].OxideCusOreGradeCrusherPercentage;
                        ws21.Cells[lastRow1, 24].Value = lstRealMovementProduction[0].OxideRecoveryCrusherAndRomPercentage;
                        ws21.Cells[lastRow1, 25].Value = lstRealMovementProduction[0].OxideRomTonnes;
                        ws21.Cells[lastRow1, 26].Value = lstRealMovementProduction[0].OxideCrushedMaterialTonnes;
                        ws21.Cells[lastRow1, 27].Value = lstRealMovementProduction[0].OxideTotalStackedMaterialTonnes;
                        ws21.Cells[lastRow1, 28].Value = lstRealMovementProduction[0].OxideCuCathodesTonnes;
                        ws21.Cells[lastRow1, 29].Value = lstRealMovementProduction[0].SulphideLeachCutOreGradePercentage;
                        ws21.Cells[lastRow1, 30].Value = lstRealMovementProduction[0].SulphideLeachCusOreGradePercentage;
                        ws21.Cells[lastRow1, 31].Value = lstRealMovementProduction[0].SulphideLeachRecoveryPercentage;
                        ws21.Cells[lastRow1, 32].Value = lstRealMovementProduction[0].SulphideLeachStackedMaterialFromMineTonnes;
                        ws21.Cells[lastRow1, 33].Value = lstRealMovementProduction[0].SulphideLeachContractorsStackedMaterialFromStocksTonnes;
                        ws21.Cells[lastRow1, 34].Value = lstRealMovementProduction[0].SulphideLeachMelStackedMaterialFromStocksTonnes;
                        ws21.Cells[lastRow1, 35].Value = lstRealMovementProduction[0].SulphideLeachTotalStackedMaterialTonnes;
                        ws21.Cells[lastRow1, 36].Value = lstRealMovementProduction[0].SulphideLeachCuCathodesTonnes;
                        ws21.Cells[lastRow1, 37].Value = lstRealMovementProduction[0].ColosoCuExColosoTonnes;
                        ws21.Cells[lastRow1, 38].Value = lstRealMovementProduction[0].ColosoCuFromLowGradeConcentrateTonnes;

                        ExcelWorksheet ws22 = pck2.Workbook.Worksheets["RealByPitDisintegrated"];
                        int lastRow2 = ws22.Dimension.End.Row + 1;

                        ws22.Cells[lastRow2, 1].Value = newDate;
                        ws22.Cells[lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws22.Cells[lastRow2, 2].Value = lstRealPitDisintegrated[0].ExpitEsTonnes;
                        ws22.Cells[lastRow2, 3].Value = lstRealPitDisintegrated[0].ExpitEnTonnes;
                        ws22.Cells[lastRow2, 4].Value = lstRealPitDisintegrated[0].TotalExpitTonnes;
                        ws22.Cells[lastRow2, 5].Value = lstRealPitDisintegrated[0].MillRehandlingTonnes;
                        ws22.Cells[lastRow2, 6].Value = lstRealPitDisintegrated[0].OlRehandlingTonnes;
                        ws22.Cells[lastRow2, 7].Value = lstRealPitDisintegrated[0].SlRehandlingTonnes;
                        ws22.Cells[lastRow2, 8].Value = lstRealPitDisintegrated[0].OtherRehandlingTonnes;
                        ws22.Cells[lastRow2, 9].Value = lstRealPitDisintegrated[0].TotalRehandlingTonnes;
                        ws22.Cells[lastRow2, 10].Value = lstRealPitDisintegrated[0].TotalMovementTonnes;
                        ws22.Cells[lastRow2, 11].Value = lstRealPitDisintegrated[0].RehandlingTotalTonnes;
                        ws22.Cells[lastRow2, 12].Value = lstRealPitDisintegrated[0].MovementTotalTonnes;
                        ws22.Cells[lastRow2, 13].Value = lstRealPitDisintegrated[0].TotalTonnes;

                        ExcelWorksheet ws23 = pck2.Workbook.Worksheets["RealMineComplianceMillFc"];
                        int lastRow3 = ws23.Dimension.End.Row + 1;

                        ws23.Cells[lastRow3, 1].Value = newDate;
                        ws23.Cells[lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws23.Cells[lastRow3, 2].Value = lstRealMillFc[0].OreGradeCut;
                        ws23.Cells[lastRow3, 3].Value = lstRealMillFc[0].MillRecovery;
                        ws23.Cells[lastRow3, 4].Value = lstRealMillFc[0].MillFeed;


                        ExcelWorksheet ws24 = pck2.Workbook.Worksheets["RealLoadingFc"];
                        int lastRow4 = ws24.Dimension.End.Row + 1;

                        ExcelWorksheet ws25 = pck2.Workbook.Worksheets["Loading SS YTD"];
                        int lastRow5 = ws25.Dimension.End.Row + 1;

                        for (int i = 0; i < lstRealLoadingFc.Count; i++)
                        {
                            ws24.Cells[i + lastRow4, 1].Value = newDate;
                            ws24.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws24.Cells[i + lastRow4, 2].Value = lstRealLoadingFc[i].Name;
                            ws24.Cells[i + lastRow4, 3].Value = lstRealLoadingFc[i].Units;
                            ws24.Cells[i + lastRow4, 4].Value = lstRealLoadingFc[i].AvailabilityPercentage;
                            ws24.Cells[i + lastRow4, 5].Value = lstRealLoadingFc[i].UtilizationPercentage;
                            ws24.Cells[i + lastRow4, 6].Value = lstRealLoadingFc[i].TotalHoursHours;
                            ws24.Cells[i + lastRow4, 7].Value = lstRealLoadingFc[i].AvailableHoursHours;
                            ws24.Cells[i + lastRow4, 8].Value = lstRealLoadingFc[i].EquipmentScheduledDowntimeHours;
                            ws24.Cells[i + lastRow4, 9].Value = lstRealLoadingFc[i].EquipmentNonScheduledDowntimeHours;
                            ws24.Cells[i + lastRow4, 10].Value = lstRealLoadingFc[i].ProcessScheduledDowntimeHours;
                            ws24.Cells[i + lastRow4, 11].Value = lstRealLoadingFc[i].ProcessNonScheduledDowntimeHours;
                            ws24.Cells[i + lastRow4, 12].Value = lstRealLoadingFc[i].StandByHours;
                            ws24.Cells[i + lastRow4, 13].Value = lstRealLoadingFc[i].HangTimeHours;
                            ws24.Cells[i + lastRow4, 14].Value = lstRealLoadingFc[i].ProductionTimeHours;
                            ws24.Cells[i + lastRow4, 15].Value = lstRealLoadingFc[i].PerformanceTonnesPerHour;
                            ws24.Cells[i + lastRow4, 16].Value = lstRealLoadingFc[i].TotalTonnesTonnes;

                            if (lstRealLoadingFc[i].Name == "SUMMARY SHOVEL 73 yd3")
                            {
                                ws25.Cells[lastRow5, 1].Value = newDate;
                                ws25.Cells[lastRow5, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                ws25.Cells[lastRow5, 2].Value = lstRealLoadingFc[i].Name;
                                ws25.Cells[lastRow5, 3].Value = lstRealLoadingFc[i].Units;
                                ws25.Cells[lastRow5, 4].Value = lstRealLoadingFc[i].AvailabilityPercentage;
                                ws25.Cells[lastRow5, 5].Value = lstRealLoadingFc[i].UtilizationPercentage;
                                ws25.Cells[lastRow5, 6].Value = lstRealLoadingFc[i].TotalHoursHours;
                                ws25.Cells[lastRow5, 7].Value = lstRealLoadingFc[i].AvailableHoursHours;
                                ws25.Cells[lastRow5, 8].Value = lstRealLoadingFc[i].EquipmentScheduledDowntimeHours;
                                ws25.Cells[lastRow5, 9].Value = lstRealLoadingFc[i].EquipmentNonScheduledDowntimeHours;
                                ws25.Cells[lastRow5, 10].Value = lstRealLoadingFc[i].ProcessScheduledDowntimeHours;
                                ws25.Cells[lastRow5, 11].Value = lstRealLoadingFc[i].ProcessNonScheduledDowntimeHours;
                                ws25.Cells[lastRow5, 12].Value = lstRealLoadingFc[i].StandByHours;
                                ws25.Cells[lastRow5, 13].Value = lstRealLoadingFc[i].HangTimeHours;
                                ws25.Cells[lastRow5, 14].Value = lstRealLoadingFc[i].ProductionTimeHours;
                                ws25.Cells[lastRow5, 15].Value = lstRealLoadingFc[i].PerformanceTonnesPerHour;
                                ws25.Cells[lastRow5, 16].Value = lstRealLoadingFc[i].TotalTonnesTonnes;
                            }

                        }

                        ExcelWorksheet ws26 = pck2.Workbook.Worksheets["RealHaulageFc"];
                        int lastRow6 = ws26.Dimension.End.Row + 1;

                        ExcelWorksheet ws27 = pck2.Workbook.Worksheets["Hauling TF YTD"];
                        int lastRow7 = ws27.Dimension.End.Row + 1;

                        for (int i = 0; i < lstRealHaulingFc.Count; i++)
                        {
                            ws26.Cells[i + lastRow6, 1].Value = newDate;
                            ws26.Cells[i + lastRow6, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws26.Cells[i + lastRow6, 2].Value = lstRealHaulingFc[i].Name;
                            ws26.Cells[i + lastRow6, 3].Value = lstRealHaulingFc[i].Units;
                            ws26.Cells[i + lastRow6, 4].Value = lstRealHaulingFc[i].MechanicalAvailabilityPercentage;
                            ws26.Cells[i + lastRow6, 5].Value = lstRealHaulingFc[i].PhysicalAvailabilityPercentage;
                            ws26.Cells[i + lastRow6, 6].Value = lstRealHaulingFc[i].UtilizationPercentage;
                            ws26.Cells[i + lastRow6, 7].Value = lstRealHaulingFc[i].TotalHoursHours;
                            ws26.Cells[i + lastRow6, 8].Value = lstRealHaulingFc[i].AvailableHoursHours;
                            ws26.Cells[i + lastRow6, 9].Value = lstRealHaulingFc[i].EquipmentScheduledDowntimeHours;
                            ws26.Cells[i + lastRow6, 10].Value = lstRealHaulingFc[i].EquipmentNonScheduledDowntimeHours;
                            ws26.Cells[i + lastRow6, 11].Value = lstRealHaulingFc[i].ProcessScheduledDowntimeHours;
                            ws26.Cells[i + lastRow6, 12].Value = lstRealHaulingFc[i].ProcessNonScheduledDowntimeHours;
                            ws26.Cells[i + lastRow6, 13].Value = lstRealHaulingFc[i].StandByHours;
                            ws26.Cells[i + lastRow6, 14].Value = lstRealHaulingFc[i].QueueTimeHours;
                            ws26.Cells[i + lastRow6, 15].Value = lstRealHaulingFc[i].ProductionTimeHoursHours;
                            ws26.Cells[i + lastRow6, 16].Value = lstRealHaulingFc[i].PerformanceTonnesPerHour;
                            ws26.Cells[i + lastRow6, 17].Value = lstRealHaulingFc[i].CycleTimeHours;
                            ws26.Cells[i + lastRow6, 18].Value = lstRealHaulingFc[i].TotalTonnesTonnes;

                            if (lstRealHaulingFc[i].Name == "Total Fleet")
                            {
                                ws27.Cells[lastRow7, 1].Value = newDate;
                                ws27.Cells[lastRow7, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                ws27.Cells[lastRow7, 2].Value = lstRealHaulingFc[i].Name;
                                ws27.Cells[lastRow7, 3].Value = lstRealHaulingFc[i].Units;
                                ws27.Cells[lastRow7, 4].Value = lstRealHaulingFc[i].MechanicalAvailabilityPercentage;
                                ws27.Cells[lastRow7, 5].Value = lstRealHaulingFc[i].PhysicalAvailabilityPercentage;
                                ws27.Cells[lastRow7, 6].Value = lstRealHaulingFc[i].UtilizationPercentage;
                                ws27.Cells[lastRow7, 7].Value = lstRealHaulingFc[i].TotalHoursHours;
                                ws27.Cells[lastRow7, 8].Value = lstRealHaulingFc[i].AvailableHoursHours;
                                ws27.Cells[lastRow7, 9].Value = lstRealHaulingFc[i].EquipmentScheduledDowntimeHours;
                                ws27.Cells[lastRow7, 10].Value = lstRealHaulingFc[i].EquipmentNonScheduledDowntimeHours;
                                ws27.Cells[lastRow7, 11].Value = lstRealHaulingFc[i].ProcessScheduledDowntimeHours;
                                ws27.Cells[lastRow7, 12].Value = lstRealHaulingFc[i].ProcessNonScheduledDowntimeHours;
                                ws27.Cells[lastRow7, 13].Value = lstRealHaulingFc[i].StandByHours;
                                ws27.Cells[lastRow7, 14].Value = lstRealHaulingFc[i].QueueTimeHours;
                                ws27.Cells[lastRow7, 15].Value = lstRealHaulingFc[i].ProductionTimeHoursHours;
                                ws27.Cells[lastRow7, 16].Value = lstRealHaulingFc[i].PerformanceTonnesPerHour;
                                ws27.Cells[lastRow7, 17].Value = lstRealHaulingFc[i].CycleTimeHours;
                                ws27.Cells[lastRow7, 18].Value = lstRealHaulingFc[i].TotalTonnesTonnes;
                            }
                        }

                        byte[] fileText2 = pck2.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText2);

                        UpdateA = $"Actualizado: {DateTime.Now}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Upload Error");
                    }
                }

            }

        }

        private void GenerateBudgetTemplate()
        {
            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Mine Compliance Budget Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws3 = pck.Workbook.Worksheets.Add("Budget Principal");
            var ws2 = pck.Workbook.Worksheets.Add("Budget Movement Production");
            var ws = pck.Workbook.Worksheets.Add("Budget Pit Disintegrated");

            List<string> lstCategory1 = new List<string>() { "Category", "Parameter", "Unit", "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            List<string> lstParameter1 = new List<string>() { "Expit ES", "Expit EN", "Mill", "OL", "SL", "Other", "Reh.", "Mov." };

            ws.Cells["B2"].Value = $"FY{FiscalYear}";
            ws.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws.Cells["B2:B5"].Style.Font.Bold = true;
            ws.Cells["B3:B5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B3:B5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E7E6E6"));


            ws.Cells["C2:J2"].Merge = true;
            ws.Cells["C2:J2"].Style.Font.Bold = true;
            ws.Cells["C2:J2"].Value = "Budget";
            ws.Cells["C2:J2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells["B2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C2:J2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C2:J2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#1F4E78"));
            ws.Cells[$"B2:J2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < 16; i++)
            {
                ws.Cells[$"B{2 + i}:J{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] G = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
            for (int i = G.GetLowerBound(0); i <= G.GetUpperBound(0); i++)
            {
                ws.Cells[$"{G[i]}2:{G[i]}17"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Column(2 + i).Width = 14;
            }

            ws.Column(1).Width = 3;
            ws.Cells["A6"].Value = "Month";
            ws.Cells["A6:A17"].Merge = true;
            ws.Cells["A6"].Style.Font.Bold = true;
            ws.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A6"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws.Cells["A6"].Style.TextRotation = 90;
            ws.Cells["A6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A17"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < lstCategory1.Count; i++)
            {
                ws.Cells[3 + i, 2].Value = lstCategory1[i];

            }

            ws.Cells["C3:D3"].Merge = true;
            ws.Cells["E3:H3"].Merge = true;
            ws.Cells["I3:J3"].Merge = true;
            ws.Cells["C3:J3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells[3, 3].Value = "Movement";
            ws.Cells[3, 5].Value = "Rehandling";
            ws.Cells[3, 9].Value = "Total";

            ws.Cells["C3:J3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
            ws.Cells["E3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E3:H3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));

            for (int i = 0; i < 3; i++)
            {
                ws.Cells[$"C{3 + i}:J{3 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < lstParameter1.Count; i++)
            {
                ws.Cells[4, 3 + i].Value = lstParameter1[i];
                ws.Cells[5, 3 + i].Value = "t";
            }
            ws.Cells["C4:J4"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws.Cells["C4:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C4:J4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
            ws.Cells["E4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E4:H4"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

            ws.Cells["C5:J5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C5:J5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
            ws.Cells["E5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E5:H5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));

            
            List<string> lstZone2 = new List<string>() { "Los Colorados", "Laguna Seca", "Laguna Seca 2", "Oxide", "Sulphide Leach" };
            List<string> lstParameter21 = new List<string>() { "Ore Grade - CuT", "Mill Recovery", "Mill Feed", "Cu Ex Mill", "Runtime", "Hours" };
            List<string> lstParameter22 = new List<string>() { "Ore to OL", "Cu Cathodes", "Stacked Material from Mine", "Contractors Stacked Material from Stocks", "MEL Stacked Material from Stocks", "Total Stacked Material", "Cu Cathodes" };
            List<string> lstUnit21 = new List<string>() { "%", "%", "t", "t Cu", "%", "h" };
            List<string> lstUnit22 = new List<string>() { "dmt", "t Cu", "dmt", "dmt", "dmt", "dmt", "t Cu" };
            List<string> lstMonth2 = new List<string>() { "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };

            ws2.Column(1).Width = 20;
            ws2.Cells["A1"].Value = "Budget";
            ws2.Row(1).Style.Font.Bold = true;
            ws2.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws2.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["A1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#002060"));
            ws2.Cells["A1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws2.Cells["D1"].Value = $"FY{FiscalYear}";
            ws2.Cells["D1:O1"].Merge = true;

            ws2.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws2.Cells["D1:O1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["D1:O1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
            ws2.Cells["D1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws2.Cells["O1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws2.Cells["D1:O1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            ws2.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            for (int i = 0; i < lstMonth2.Count; i++)
            {
                ws2.Cells[2, 4 + i].Value = lstMonth2[i];
                ws2.Column(4 + i).Width = 11;
            }
            ws2.Cells["D2:O2"].Style.Font.Bold = true;
            ws2.Cells["D2:O2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws2.Cells["D2:O2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["D2:O2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            int[] H = { 3, 11, 19, 27, 29 };
            int[] HH = { 8, 16, 24, 28, 33 };
            for (int i = H.GetLowerBound(0); i <= H.GetUpperBound(0); i++)
            {
                ws2.Cells[H[i], 1].Value = lstZone2[i];
                ws2.Cells[$"A{H[i]}:O{H[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws2.Cells[H[i], 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[H[i], 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws2.Cells[$"B{H[i]}:B{HH[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                if (i % 2 != 0)
                {
                    ws2.Cells[H[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[H[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));
                }
                else if (i % 2 == 0)
                {
                    ws2.Cells[H[i], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[H[i], 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
                }
            }

            ws2.Column(2).Width = 35;
            for (int i = 0; i < lstParameter21.Count; i++)
            {
                ws2.Cells[3 + i, 2].Value = lstParameter21[i];
                ws2.Cells[3 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{3 + i}:O{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ws2.Cells[11 + i, 2].Value = lstParameter21[i];
                ws2.Cells[11 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{11 + i}:O{11 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ws2.Cells[19 + i, 2].Value = lstParameter21[i];
                ws2.Cells[19 + i, 3].Value = lstUnit21[i];
                ws2.Cells[$"B{19 + i}:O{19 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] I = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O" };
            for (int i = I.GetLowerBound(0); i <= I.GetUpperBound(0); i++)
            {
                ws2.Cells[$"{I[i]}3:{I[i]}8"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"{I[i]}11:{I[i]}16"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"{I[i]}19:{I[i]}24"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"{I[i]}27:{I[i]}33"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }


            int[] J = { 3, 19, 29 };
            int[] K = { 8, 24, 33 };
            for (int i = J.GetLowerBound(0); i <= J.GetUpperBound(0); i++)
            {
                ws2.Cells[$"B{J[i]}:B{K[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"B{J[i]}:B{K[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                ws2.Cells[$"C{J[i]}:C{K[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"C{J[i]}:C{K[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
            }

            int[] L = { 11, 27 };
            int[] M = { 16, 28 };
            for (int i = L.GetLowerBound(0); i <= L.GetUpperBound(0); i++)
            {
                ws2.Cells[$"B{L[i]}:B{M[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"B{L[i]}:B{M[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                ws2.Cells[$"C{L[i]}:C{M[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[$"C{L[i]}:C{M[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
            }

            ws2.Column(3).Width = 13;
            ws2.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int[] N = { 2, 10, 18, 26 };
            for (int i = N.GetLowerBound(0); i <= N.GetUpperBound(0); i++)
            {
                ws2.Cells[N[i], 3].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws2.Cells[N[i], 3].Value = "Unit";
                ws2.Cells[N[i], 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws2.Cells[N[i], 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[N[i], 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                if (i % 2 != 0)
                {
                    ws2.Cells[N[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[N[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));
                }
                else if (i % 2 == 0)
                {
                    ws2.Cells[N[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws2.Cells[N[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                }
            }

            for (int i = 0; i < lstUnit22.Count; i++)
            {
                ws2.Cells[27 + i, 2].Value = lstParameter22[i];
                ws2.Cells[27 + i, 3].Value = lstUnit22[i];

                ws2.Cells[$"B{27 + i}:O{27 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }


            
            List<string> lstParameter3 = new List<string>() { "Expit", "Rehandling", "Movement", "Shovels units 73 yd3", "Shovels availability", "Shovels utilization", "Shovels performance", "Stand By", "Production time", "Available hours", "Shovel hours", "Trucks units", "Trucks availability", "Trucks utilization", "Trucks performance", "Stand By", "Trucks hours", "Production Time", "Available hours", "Mill Throughput", "Mill Grade", "Mill Recovery", "Mill Rehandle", "OL Throughput", "OL Grade", "Recovery", "CuS", "SL Throughput", "SL Grade", "Recovery", "CuS", "Mill Production", "Cathodes", "Total Production" };
            List<string> lstUnit3 = new List<string>() { "Mt", "Mt", "Mt", "eq", "%", "%", "t/h", "h", "h", "h", "h", "eq", "%", "%", "t/h", "h", "h", "h", "h", "Mt", "Cu %", "%", "%", "Mt", "Cu %", "%", "%", "Mt", "Cu %", "%", "%", "kt", "kt", "kt" };
            List<string> lstMonth3 = new List<string>() { "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };

            ws3.Column(2).Width = 20;
            ws3.Column(2).Style.Font.Bold = true;
            ws3.Cells["B2"].Value = "BUDGET B01";
            ws3.Row(2).Style.Font.Bold = true;
            ws3.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#1F4E78"));
            ws3.Cells["B2:B37"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

            ws3.Cells["D2"].Value = $"FY{FiscalYear}";
            ws3.Cells["D2:O2"].Merge = true;

            ws3.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws3.Cells["D2:O2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["D2:O2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
            ws3.Cells["D2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws3.Cells["C3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws3.Cells["O2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws3.Cells["D2:O2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            ws3.Cells["D2:O2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            ws3.Cells["B3:O3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws3.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Row(3).Style.Font.Bold = true;

            for (int i = 0; i < lstMonth3.Count; i++)
            {
                ws3.Cells[3, 4 + i].Value = lstMonth3[i];
                ws3.Column(4 + i).Width = 11;
            }
            ws3.Cells["D3:O3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws3.Cells["D3:O3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["D3:O3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            ws3.Column(3).Width = 10;
            ws3.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < lstParameter3.Count; i++)
            {
                ws3.Cells[4 + i, 2].Value = lstParameter3[i];
                ws3.Cells[4 + i, 3].Value = lstUnit3[i];
                ws3.Cells[$"B{4 + i}:O{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] O = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O" };
            for (int i = O.GetLowerBound(0); i <= O.GetUpperBound(0); i++)
            {
                ws3.Cells[$"{I[i]}4:{I[i]}37"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }


            int[] P = { 4, 15, 27, 35 };
            int[] Q = { 6, 22, 30, 37 };
            for (int i = P.GetLowerBound(0); i <= P.GetUpperBound(0); i++)
            {
                ws3.Cells[$"B{P[i]}:B{Q[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[$"B{P[i]}:B{Q[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                ws3.Cells[$"C{P[i]}:C{Q[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[$"C{P[i]}:C{Q[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws3.Cells[$"D{Q[i]}:O{Q[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            }

            int[] R = { 7, 23, 31 };
            int[] S = { 14, 26, 34 };
            for (int i = R.GetLowerBound(0); i <= R.GetUpperBound(0); i++)
            {
                ws3.Cells[$"B{R[i]}:B{S[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[$"B{R[i]}:B{S[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                ws3.Cells[$"C{R[i]}:C{S[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[$"C{R[i]}:C{S[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                ws3.Cells[$"D{S[i]}:O{S[i]}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "MineComplianceBudgetTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                FileStream fs = File.OpenWrite(dialog.FileName);
                fs.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabled2 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Upload Error");
            }
        }

        public class BudgetPrincipal 
        {
            public DateTime Date { get; set; }
            public double ExpitTonnes { get; set; }
            public double RehandlingTonnes { get; set; }
            public double MovementTonnes { get; set; }

            public double ShovelsUnits73Yd3 { get; set; }
            public double ShovelsAvailabilityPercentage { get; set; }
            public double ShovelsUtilizationPercentage { get; set; }
            public double ShovelsPerformanceTonnesPerHour { get; set; }
            public double ShovelsStandByHours { get; set; }
            public double ShovelsProductionTimeHours { get; set; }
            public double ShovelAvailableHoursHours { get; set; }
            public double ShovelHoursHours { get; set; }

            public double TrucksUnits { get; set; }
            public double TrucksAvailabilityPercentage { get; set; }
            public double TrucksUtilizationPercentage { get; set; }
            public double TrucksPerformanceTonnesPerDay { get; set; }
            public double TrucksStandByHours { get; set; }
            public double TrucksHoursHours { get; set; }
            public double TrucksProductionTimeHours { get; set; }
            public double TrucksAvailableHoursHours { get; set; }

            public double MillThroughputTonnes { get; set; }
            public double MillGradeCuPercentage { get; set; }
            public double MillRecoveryPercentage { get; set; }
            public double MillRehandlingPercentage { get; set; }

            public double OlThroughputTonnes { get; set; }
            public double OlGradeCuPercentage { get; set; }
            public double OlRecoveryPercentage { get; set; }
            public double OlCuSPercentage { get; set; }

            public double SlThroughputTonnes { get; set; }
            public double SlGradeCuPercentage { get; set; }
            public double SlRecoveryPercentage { get; set; }
            public double SlCuSPercentage { get; set; }

            public double MillProductionTonnes { get; set; }
            public double CathodesTonnes { get; set; }
            public double TotalProductionTonnes { get; set; }
        }

        readonly List<BudgetPrincipal> lstBudgetPrincipal = new List<BudgetPrincipal>();
        
        public class BudgetMovementProduction
        {
            public DateTime Date { get; set; }
            public double LosColoradosOreGradeCutPercentage { get; set; }
            public double LosColoradosMillRecoveryPercentage { get; set; }
            public double LosColoradosMillFeedTonnes { get; set; }
            public double LosColoradosCuExMillTonnes { get; set; }
            public double LosColoradosRuntimePercentage { get; set; }
            public double LosColoradosHoursHours { get; set; }

            public double LagunaSecaOreGradeCutPercentage { get; set; }
            public double LagunaSecaMillRecoveryPercentage { get; set; }
            public double LagunaSecaMillFeedTonnes { get; set; }
            public double LagunaSecaCuExMillTonnes { get; set; }
            public double LagunaSecaRuntimePercentage { get; set; }
            public double LagunaSecaHoursHours { get; set; }

            public double LagunaSeca2OreGradeCutPercentage { get; set; }
            public double LagunaSeca2MillRecoveryPercentage { get; set; }
            public double LagunaSeca2MillFeedTonnes { get; set; }
            public double LagunaSeca2CuExMillTonnes { get; set; }
            public double LagunaSeca2RuntimePercentage { get; set; }
            public double LagunaSeca2HoursHours { get; set; }

            public double OxideOreToOlTonnes { get; set; }
            public double OxideCuCathodesTonnes { get; set; }

            public double SulphideLeachStackedMaterialFromMineTonnes { get; set; }
            public double SulphideLeachContractorsStackedMaterialFromStocksTonnes { get; set; }
            public double SulphideLeachMelStackedMaterialFromStocksTonnesTonnes { get; set; }
            public double SulphideLeachTotalStackedMaterialTonnes { get; set; }
            public double SulphideLeachCuCathodesTonnes { get; set; }

        }

        readonly List<BudgetMovementProduction> lstBudgetMovementProduction = new List<BudgetMovementProduction>();

        public class BudgetPitDisintegrated
        {
            public DateTime Date { get; set; }
            public double ExpitEsTonnes { get; set; }
            public double ExpitEnTonnes { get; set; }
            public double TotalExpitTonnes { get; set; }

            public double MillRehandlingTonnes { get; set; }
            public double OlRehandlingTonnes { get; set; }
            public double SlRehandlingTonnes { get; set; }
            public double OtherRehandlingTonnes { get; set; }
            public double TotalRehandlingTonnes { get; set; }

            public double TotalMovementTonnes { get; set; }

            public double RehandlingTotalTonnes { get; set; }
            public double MovementTotalTonnes { get; set; }

            public double TotalTonnes { get; set; }

        }

        readonly List<BudgetPitDisintegrated> lstBudgetPitDisintegrated = new List<BudgetPitDisintegrated>();


        private void LoadBudgetTemplate()
        {
            lstBudgetPrincipal.Clear();
            lstBudgetMovementProduction.Clear();
            lstBudgetPitDisintegrated.Clear();

            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (op.ShowDialog() == true)
            {

                FileInfo FilePath = new FileInfo(op.FileName);
                ExcelPackage pck = new ExcelPackage(FilePath);

                try
                {
                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    ExcelWorksheet ws = pck.Workbook.Worksheets["Budget Principal"];
                    ExcelWorksheet ws2 = pck.Workbook.Worksheets["Budget Movement Production"];
                    ExcelWorksheet ws3 = pck.Workbook.Worksheets["Budget Pit Disintegrated"];                              

                    DateTime Db = DateTime.Now;

                    for (int i = 0; i < 12; i++)
                    {
                        int M = DateTime.ParseExact(ws.Cells[3, 4 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;

                        if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i == 5)
                        {
                            Db = new DateTime(FiscalYear - 1, M, 1, 00, 00, 00).AddMilliseconds(000);
                        }
                        else if (i == 6 || i == 7 || i == 8 || i == 9 || i == 10 || i == 11)
                        {
                            Db = new DateTime(FiscalYear, M, 1, 00, 00, 00).AddMilliseconds(000);
                        }

                        for (int j = 0; j < 34; j++)
                        {
                            if (ws.Cells[4 + j, 4 + i].Value == null)
                            {
                                ws.Cells[4 + j, 4 + i].Value = 0;
                            }

                            if (ws2.Cells[3 + j, 4 + i].Value == null)
                            {
                                ws2.Cells[3 + j, 4 + i].Value = 0;
                            }

                        }

                        lstBudgetPrincipal.Add(new BudgetPrincipal()
                        {
                            Date = Db,
                            ExpitTonnes = double.Parse(ws.Cells[4, 4 + i].Value.ToString()),
                            RehandlingTonnes = double.Parse(ws.Cells[5, 4 + i].Value.ToString()),
                            MovementTonnes = double.Parse(ws.Cells[6, 4 + i].Value.ToString()),
                            ShovelsUnits73Yd3 = double.Parse(ws.Cells[7, 4 + i].Value.ToString()),
                            ShovelsAvailabilityPercentage = double.Parse(ws.Cells[8, 4 + i].Value.ToString()),
                            ShovelsUtilizationPercentage = double.Parse(ws.Cells[9, 4 + i].Value.ToString()),
                            ShovelsPerformanceTonnesPerHour = double.Parse(ws.Cells[10, 4 + i].Value.ToString())/1000000,
                            ShovelsStandByHours = double.Parse(ws.Cells[11, 4 + i].Value.ToString()),
                            ShovelsProductionTimeHours = double.Parse(ws.Cells[12, 4 + i].Value.ToString()),
                            ShovelAvailableHoursHours = double.Parse(ws.Cells[13, 4 + i].Value.ToString()),
                            ShovelHoursHours = double.Parse(ws.Cells[14, 4 + i].Value.ToString()),
                            TrucksUnits = double.Parse(ws.Cells[15, 4 + i].Value.ToString()),
                            TrucksAvailabilityPercentage = double.Parse(ws.Cells[16, 4 + i].Value.ToString()),
                            TrucksUtilizationPercentage = double.Parse(ws.Cells[17, 4 + i].Value.ToString()),
                            TrucksPerformanceTonnesPerDay = double.Parse(ws.Cells[18, 4 + i].Value.ToString())/1000000,
                            TrucksStandByHours = double.Parse(ws.Cells[19, 4 + i].Value.ToString()),
                            TrucksHoursHours = double.Parse(ws.Cells[20, 4 + i].Value.ToString()),
                            TrucksProductionTimeHours = double.Parse(ws.Cells[21, 4 + i].Value.ToString()),
                            TrucksAvailableHoursHours = double.Parse(ws.Cells[22, 4 + i].Value.ToString()),
                            MillThroughputTonnes = double.Parse(ws.Cells[23, 4 + i].Value.ToString()),
                            MillGradeCuPercentage = double.Parse(ws.Cells[24, 4 + i].Value.ToString()),
                            MillRecoveryPercentage = double.Parse(ws.Cells[25, 4 + i].Value.ToString()),
                            MillRehandlingPercentage = double.Parse(ws.Cells[26, 4 + i].Value.ToString()),
                            OlThroughputTonnes = double.Parse(ws.Cells[27, 4 + i].Value.ToString()),
                            OlGradeCuPercentage = double.Parse(ws.Cells[28, 4 + i].Value.ToString()),
                            OlRecoveryPercentage = double.Parse(ws.Cells[29, 4 + i].Value.ToString()),
                            OlCuSPercentage = double.Parse(ws.Cells[30, 4 + i].Value.ToString()),
                            SlThroughputTonnes = double.Parse(ws.Cells[31, 4 + i].Value.ToString()),
                            SlGradeCuPercentage = double.Parse(ws.Cells[32, 4 + i].Value.ToString()),
                            SlRecoveryPercentage = double.Parse(ws.Cells[33, 4 + i].Value.ToString()),
                            SlCuSPercentage = double.Parse(ws.Cells[34, 4 + i].Value.ToString()),
                            MillProductionTonnes = double.Parse(ws.Cells[35, 4 + i].Value.ToString()),
                            CathodesTonnes = double.Parse(ws.Cells[36, 4 + i].Value.ToString()),
                            TotalProductionTonnes = double.Parse(ws.Cells[37, 4 + i].Value.ToString())
                        });

                        lstBudgetMovementProduction.Add(new BudgetMovementProduction()
                        {
                            Date = Db,
                            LosColoradosOreGradeCutPercentage = double.Parse(ws2.Cells[3, 4 + i].Value.ToString()),
                            LosColoradosMillRecoveryPercentage = double.Parse(ws2.Cells[4, 4 + i].Value.ToString()),
                            LosColoradosMillFeedTonnes = double.Parse(ws2.Cells[5, 4 + i].Value.ToString())/1000,
                            LosColoradosCuExMillTonnes = double.Parse(ws2.Cells[6, 4 + i].Value.ToString())/1000,
                            LosColoradosRuntimePercentage = double.Parse(ws2.Cells[7, 4 + i].Value.ToString()),
                            LosColoradosHoursHours = double.Parse(ws2.Cells[8, 4 + i].Value.ToString()),
                            LagunaSecaOreGradeCutPercentage = double.Parse(ws2.Cells[11, 4 + i].Value.ToString()),
                            LagunaSecaMillRecoveryPercentage = double.Parse(ws2.Cells[12, 4 + i].Value.ToString()),
                            LagunaSecaMillFeedTonnes = double.Parse(ws2.Cells[13, 4 + i].Value.ToString())/1000,
                            LagunaSecaCuExMillTonnes = double.Parse(ws2.Cells[14, 4 + i].Value.ToString())/1000,
                            LagunaSecaRuntimePercentage = double.Parse(ws2.Cells[15, 4 + i].Value.ToString()),
                            LagunaSecaHoursHours = double.Parse(ws2.Cells[16, 4 + i].Value.ToString()),
                            LagunaSeca2OreGradeCutPercentage = double.Parse(ws2.Cells[19, 4 + i].Value.ToString()),
                            LagunaSeca2MillRecoveryPercentage = double.Parse(ws2.Cells[20, 4 + i].Value.ToString()),
                            LagunaSeca2MillFeedTonnes = double.Parse(ws2.Cells[21, 4 + i].Value.ToString())/1000,
                            LagunaSeca2CuExMillTonnes = double.Parse(ws2.Cells[22, 4 + i].Value.ToString())/1000,
                            LagunaSeca2RuntimePercentage = double.Parse(ws2.Cells[23, 4 + i].Value.ToString()),
                            LagunaSeca2HoursHours = double.Parse(ws2.Cells[24, 4 + i].Value.ToString()),
                            OxideOreToOlTonnes = double.Parse(ws2.Cells[27, 4 + i].Value.ToString())/1000,
                            OxideCuCathodesTonnes = double.Parse(ws2.Cells[28, 4 + i].Value.ToString())/1000,
                            SulphideLeachStackedMaterialFromMineTonnes = double.Parse(ws2.Cells[29, 4 + i].Value.ToString())/1000,
                            SulphideLeachContractorsStackedMaterialFromStocksTonnes = double.Parse(ws2.Cells[30, 4 + i].Value.ToString())/1000,
                            SulphideLeachMelStackedMaterialFromStocksTonnesTonnes = double.Parse(ws2.Cells[31, 4 + i].Value.ToString())/1000,
                            SulphideLeachTotalStackedMaterialTonnes = double.Parse(ws2.Cells[32, 4 + i].Value.ToString())/1000,
                            SulphideLeachCuCathodesTonnes = double.Parse(ws2.Cells[33, 4 + i].Value.ToString())/1000
                        });

                        lstBudgetPitDisintegrated.Add(new BudgetPitDisintegrated()
                        {
                            Date = Db,
                            ExpitEsTonnes = double.Parse(ws3.Cells[6 + i, 3].Value.ToString()) / 1000000,
                            ExpitEnTonnes = double.Parse(ws3.Cells[6 + i, 4].Value.ToString()) / 1000000,
                            TotalExpitTonnes = (double.Parse(ws3.Cells[6 + i, 3].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 4].Value.ToString())) / 1000000,                           
                            MillRehandlingTonnes = double.Parse(ws3.Cells[6 + i, 5].Value.ToString()) / 1000000,
                            OlRehandlingTonnes = double.Parse(ws3.Cells[6 + i, 6].Value.ToString()) / 1000000,
                            SlRehandlingTonnes = double.Parse(ws3.Cells[6 + i, 7].Value.ToString()) / 1000000,
                            OtherRehandlingTonnes = double.Parse(ws3.Cells[6 + i, 8].Value.ToString()) / 1000000,
                            TotalRehandlingTonnes = (double.Parse(ws3.Cells[6 + i, 5].Value.ToString())+ double.Parse(ws3.Cells[6 + i, 6].Value.ToString())+ double.Parse(ws3.Cells[6 + i, 7].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 8].Value.ToString())) / 1000000,
                            TotalMovementTonnes = (double.Parse(ws3.Cells[6 + i, 3].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 4].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 5].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 6].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 7].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 8].Value.ToString())) / 1000000,
                            RehandlingTotalTonnes = double.Parse(ws3.Cells[6 + i, 9].Value.ToString()) / 1000000,
                            MovementTotalTonnes = double.Parse(ws3.Cells[6 + i, 10].Value.ToString()) / 1000000,
                            TotalTonnes = (double.Parse(ws3.Cells[6 + i, 9].Value.ToString()) + double.Parse(ws3.Cells[6 + i, 10].Value.ToString())) / 1000000
                        });                    

                    }

                    pck.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\MineComplianceData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        ExcelWorksheet ws21 = pck2.Workbook.Worksheets["BudgetPrincipal"];
                        int lastRow1 = ws21.Dimension.End.Row + 1;

                        for (int i = 0; i < lstBudgetPrincipal.Count; i++)
                        {
                            ws21.Cells[i + lastRow1, 1].Value = lstBudgetPrincipal[i].Date;
                            ws21.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws21.Cells[i + lastRow1, 2].Value = lstBudgetPrincipal[i].ExpitTonnes;
                            ws21.Cells[i + lastRow1, 3].Value = lstBudgetPrincipal[i].RehandlingTonnes;
                            ws21.Cells[i + lastRow1, 4].Value = lstBudgetPrincipal[i].MovementTonnes;
                            ws21.Cells[i + lastRow1, 5].Value = lstBudgetPrincipal[i].ShovelsUnits73Yd3;
                            ws21.Cells[i + lastRow1, 6].Value = lstBudgetPrincipal[i].ShovelsAvailabilityPercentage;
                            ws21.Cells[i + lastRow1, 7].Value = lstBudgetPrincipal[i].ShovelsUtilizationPercentage;
                            ws21.Cells[i + lastRow1, 8].Value = lstBudgetPrincipal[i].ShovelsPerformanceTonnesPerHour;
                            ws21.Cells[i + lastRow1, 9].Value = lstBudgetPrincipal[i].ShovelsStandByHours;
                            ws21.Cells[i + lastRow1, 10].Value = lstBudgetPrincipal[i].ShovelsProductionTimeHours;
                            ws21.Cells[i + lastRow1, 11].Value = lstBudgetPrincipal[i].ShovelAvailableHoursHours;
                            ws21.Cells[i + lastRow1, 12].Value = lstBudgetPrincipal[i].ShovelHoursHours;
                            ws21.Cells[i + lastRow1, 13].Value = lstBudgetPrincipal[i].TrucksUnits;
                            ws21.Cells[i + lastRow1, 14].Value = lstBudgetPrincipal[i].TrucksAvailabilityPercentage;
                            ws21.Cells[i + lastRow1, 15].Value = lstBudgetPrincipal[i].TrucksUtilizationPercentage;
                            ws21.Cells[i + lastRow1, 16].Value = lstBudgetPrincipal[i].TrucksPerformanceTonnesPerDay;
                            ws21.Cells[i + lastRow1, 17].Value = lstBudgetPrincipal[i].TrucksStandByHours;
                            ws21.Cells[i + lastRow1, 18].Value = lstBudgetPrincipal[i].TrucksHoursHours;
                            ws21.Cells[i + lastRow1, 19].Value = lstBudgetPrincipal[i].TrucksProductionTimeHours;
                            ws21.Cells[i + lastRow1, 20].Value = lstBudgetPrincipal[i].TrucksAvailableHoursHours;
                            ws21.Cells[i + lastRow1, 21].Value = lstBudgetPrincipal[i].MillThroughputTonnes;
                            ws21.Cells[i + lastRow1, 22].Value = lstBudgetPrincipal[i].MillGradeCuPercentage;
                            ws21.Cells[i + lastRow1, 23].Value = lstBudgetPrincipal[i].MillRecoveryPercentage;
                            ws21.Cells[i + lastRow1, 24].Value = lstBudgetPrincipal[i].MillRehandlingPercentage;
                            ws21.Cells[i + lastRow1, 25].Value = lstBudgetPrincipal[i].OlThroughputTonnes;
                            ws21.Cells[i + lastRow1, 26].Value = lstBudgetPrincipal[i].OlGradeCuPercentage;
                            ws21.Cells[i + lastRow1, 27].Value = lstBudgetPrincipal[i].OlRecoveryPercentage;
                            ws21.Cells[i + lastRow1, 28].Value = lstBudgetPrincipal[i].OlCuSPercentage;
                            ws21.Cells[i + lastRow1, 29].Value = lstBudgetPrincipal[i].SlThroughputTonnes;
                            ws21.Cells[i + lastRow1, 30].Value = lstBudgetPrincipal[i].SlGradeCuPercentage;
                            ws21.Cells[i + lastRow1, 31].Value = lstBudgetPrincipal[i].SlRecoveryPercentage;
                            ws21.Cells[i + lastRow1, 32].Value = lstBudgetPrincipal[i].SlCuSPercentage;
                            ws21.Cells[i + lastRow1, 33].Value = lstBudgetPrincipal[i].MillProductionTonnes;
                            ws21.Cells[i + lastRow1, 34].Value = lstBudgetPrincipal[i].CathodesTonnes;
                            ws21.Cells[i + lastRow1, 35].Value = lstBudgetPrincipal[i].TotalProductionTonnes;
                        }

                        ExcelWorksheet ws22 = pck2.Workbook.Worksheets["BudgetMovementProduction"];
                        int lastRow2 = ws22.Dimension.End.Row + 1;

                        for (int i = 0; i < lstBudgetMovementProduction.Count; i++)
                        {
                            ws22.Cells[i + lastRow2, 1].Value = lstBudgetMovementProduction[i].Date;
                            ws22.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws22.Cells[i + lastRow2, 2].Value = lstBudgetMovementProduction[i].LosColoradosOreGradeCutPercentage;
                            ws22.Cells[i + lastRow2, 3].Value = lstBudgetMovementProduction[i].LosColoradosMillRecoveryPercentage;
                            ws22.Cells[i + lastRow2, 4].Value = lstBudgetMovementProduction[i].LosColoradosMillFeedTonnes;
                            ws22.Cells[i + lastRow2, 5].Value = lstBudgetMovementProduction[i].LosColoradosCuExMillTonnes;
                            ws22.Cells[i + lastRow2, 6].Value = lstBudgetMovementProduction[i].LosColoradosRuntimePercentage;
                            ws22.Cells[i + lastRow2, 7].Value = lstBudgetMovementProduction[i].LosColoradosHoursHours;
                            ws22.Cells[i + lastRow2, 8].Value = lstBudgetMovementProduction[i].LagunaSecaOreGradeCutPercentage;
                            ws22.Cells[i + lastRow2, 9].Value = lstBudgetMovementProduction[i].LagunaSecaMillRecoveryPercentage;
                            ws22.Cells[i + lastRow2, 10].Value = lstBudgetMovementProduction[i].LagunaSecaMillFeedTonnes;
                            ws22.Cells[i + lastRow2, 11].Value = lstBudgetMovementProduction[i].LagunaSecaCuExMillTonnes;
                            ws22.Cells[i + lastRow2, 12].Value = lstBudgetMovementProduction[i].LagunaSecaRuntimePercentage;
                            ws22.Cells[i + lastRow2, 13].Value = lstBudgetMovementProduction[i].LagunaSecaHoursHours;
                            ws22.Cells[i + lastRow2, 14].Value = lstBudgetMovementProduction[i].LagunaSeca2OreGradeCutPercentage;
                            ws22.Cells[i + lastRow2, 15].Value = lstBudgetMovementProduction[i].LagunaSeca2MillRecoveryPercentage;
                            ws22.Cells[i + lastRow2, 16].Value = lstBudgetMovementProduction[i].LagunaSeca2MillFeedTonnes;
                            ws22.Cells[i + lastRow2, 17].Value = lstBudgetMovementProduction[i].LagunaSeca2CuExMillTonnes;
                            ws22.Cells[i + lastRow2, 18].Value = lstBudgetMovementProduction[i].LagunaSeca2RuntimePercentage;
                            ws22.Cells[i + lastRow2, 19].Value = lstBudgetMovementProduction[i].LagunaSeca2HoursHours;
                            ws22.Cells[i + lastRow2, 20].Value = lstBudgetMovementProduction[i].OxideOreToOlTonnes;
                            ws22.Cells[i + lastRow2, 21].Value = lstBudgetMovementProduction[i].OxideCuCathodesTonnes;
                            ws22.Cells[i + lastRow2, 22].Value = lstBudgetMovementProduction[i].SulphideLeachStackedMaterialFromMineTonnes;
                            ws22.Cells[i + lastRow2, 23].Value = lstBudgetMovementProduction[i].SulphideLeachContractorsStackedMaterialFromStocksTonnes;
                            ws22.Cells[i + lastRow2, 24].Value = lstBudgetMovementProduction[i].SulphideLeachMelStackedMaterialFromStocksTonnesTonnes;
                            ws22.Cells[i + lastRow2, 25].Value = lstBudgetMovementProduction[i].SulphideLeachTotalStackedMaterialTonnes;
                            ws22.Cells[i + lastRow2, 26].Value = lstBudgetMovementProduction[i].SulphideLeachCuCathodesTonnes;
                        }

                        ExcelWorksheet ws23 = pck2.Workbook.Worksheets["BudgetByPitDisintegrated"];
                        int lastRow3 = ws23.Dimension.End.Row + 1;

                        for (int i = 0; i < lstBudgetPitDisintegrated.Count; i++)
                        {
                            ws23.Cells[i + lastRow3, 1].Value = lstBudgetPitDisintegrated[i].Date;
                            ws23.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws23.Cells[i + lastRow3, 2].Value = lstBudgetPitDisintegrated[i].ExpitEsTonnes;
                            ws23.Cells[i + lastRow3, 3].Value = lstBudgetPitDisintegrated[i].ExpitEnTonnes;
                            ws23.Cells[i + lastRow3, 4].Value = lstBudgetPitDisintegrated[i].TotalExpitTonnes;
                            ws23.Cells[i + lastRow3, 5].Value = lstBudgetPitDisintegrated[i].MillRehandlingTonnes;
                            ws23.Cells[i + lastRow3, 6].Value = lstBudgetPitDisintegrated[i].OlRehandlingTonnes;
                            ws23.Cells[i + lastRow3, 7].Value = lstBudgetPitDisintegrated[i].SlRehandlingTonnes;
                            ws23.Cells[i + lastRow3, 8].Value = lstBudgetPitDisintegrated[i].OtherRehandlingTonnes;
                            ws23.Cells[i + lastRow3, 9].Value = lstBudgetPitDisintegrated[i].TotalRehandlingTonnes;
                            ws23.Cells[i + lastRow3, 10].Value = lstBudgetPitDisintegrated[i].TotalMovementTonnes;
                            ws23.Cells[i + lastRow3, 11].Value = lstBudgetPitDisintegrated[i].RehandlingTotalTonnes;
                            ws23.Cells[i + lastRow3, 12].Value = lstBudgetPitDisintegrated[i].MovementTotalTonnes;
                            ws23.Cells[i + lastRow3, 13].Value = lstBudgetPitDisintegrated[i].TotalTonnes;

                        }

                        byte[] fileText2 = pck2.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText2);

                        UpdateB = $"Actualizado: {DateTime.Now}";

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Upload Error");
                    }
                }


               





            }
        }
    }
}
