using BhpAssetComplianceWpfOneDesktop.Models;
using BhpAssetComplianceWpfOneDesktop.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prism.Mvvm;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Drawing;
using OfficeOpenXml.Style;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class QuartersReconciliationFactorsViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.QuartersReconciliationFactors;

        private string _UpdateA;
        public string UpdateA
        {
            get { return _UpdateA; }
            set { SetProperty(ref _UpdateA, value); }
        }

        DateTime _Date;
        public DateTime Date
        {
            get { return _Date; }
            set { SetProperty(ref _Date, value); }
        }

        private bool _isEnabled1;
        public bool IsEnabled1
        {
            get { return _isEnabled1; }
            set { SetProperty(ref _isEnabled1, value); }
        }

        public DelegateCommand GenerarT { get; private set; }
        public DelegateCommand CargarT { get; private set; }

        public QuartersReconciliationFactorsViewModel()
        {
            Date = DateTime.Now;
            IsEnabled1 = false;
            GenerarT = new DelegateCommand(GenerateTemplate);
            CargarT = new DelegateCommand(LoadTemplate).ObservesCanExecute(() => IsEnabled1);
        }

        private void GenerateTemplate()
        {
            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Reconciliation Factors Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("Rolling Twelve Months");

            List<string> lstHeader = new List<string>() { "F0", "F1", "F2", "F3" };
            List<string> lstSecond = new List<string>() { "Ore", "%CuT", "Cu Fines", "Ore", "%CuT", "Cu Fines", "Ore", "%CuT", "Cu Fines", "Cu Fines" };
            List<string> lstColumn = new List<string>() { "Mill", "Q1", "Q2", "Q3", "Q4", "OL", "Q1", "Q2", "Q3", "Q4", "SL", "Q1", "Q2", "Q3", "Q4" };

            ws.Column(2).Width = 11;
            int[] c2 = { 7, 18, 29 };
            int[] c22 = { 10, 21, 32 };
            string FY = $"{Date.Year}";
            int Year = Int32.Parse(FY.Substring((FY.Length - 2), 2));
            if (Date.Month == 7 || Date.Month == 8 || Date.Month == 9 || Date.Month == 10 || Date.Month == 11 || Date.Month == 12)
            {
                Year = Year + 1;
            }
            string date = $"FY{Year}";

            for (int i = c2.GetLowerBound(0); i <= c2.GetUpperBound(0); i++)
            {
                ws.Cells[$"B{c2[i]}:B{c22[i]}"].Merge = true;
                ws.Cells[c2[i], 2].Style.Font.Bold = true;
                ws.Cells[c2[i], 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[c2[i], 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
                ws.Cells[c2[i], 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[c2[i], 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[$"B{c2[i]}"].Value = date;
                ws.Cells[$"B{c2[i]}:B{c2[i] + 3}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                ws.Cells[$"B{c2[i]}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            }

            ws.Column(3).Width = 13;
            ws.Column(3).Style.Font.Bold = true;
            ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Column(3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B10:M10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells["B21:L21"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells["B32:I32"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells["C3:M3"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells["C14:L14"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells["C25:I25"].Style.Border.Top.Style = ExcelBorderStyle.Thick;


            int[] c3 = { 3, 14, 25 };
            int[] c33 = { 6, 17, 28 };
            for (int i = c3.GetLowerBound(0); i <= c3.GetUpperBound(0); i++)
            {
                ws.Cells[$"C{c3[i]}:C{c33[i]}"].Merge = true;
                ws.Cells[c3[i], 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[c3[i], 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
                ws.Cells[$"D{7 + i}:M{7 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"D{18 + i}:L{18 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"D{29 + i}:I{29 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            ws.Cells["C3"].Value = lstColumn[0];
            ws.Cells["C14"].Value = lstColumn[5];
            ws.Cells["C25"].Value = lstColumn[10];
            string[] D = { "D", "E", "F", "G", "H" };
            for (int i = 0; i < 5; i++)
            {
                ws.Cells[6 + i, 3].Value = lstColumn[i];
                ws.Cells[17 + i, 3].Value = lstColumn[i + 5];
                ws.Cells[28 + i, 3].Value = lstColumn[i + 10];
                ws.Cells[$"{D[i]}25:{D[i]}32"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            string[] F = { "D", "E", "F", "G", "H", "I", "J", "K" };
            for (int i = F.GetLowerBound(0); i <= F.GetUpperBound(0); i++)
            {
                ws.Cells[$"{F[i]}14:{F[i]}21"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            string[] r = { "B", "C" };

            for (int i = r.GetLowerBound(0); i <= r.GetUpperBound(0); i++)
            {
                ws.Cells[$"{r[i]}3:{r[i]}10"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                ws.Cells[$"{r[i]}14:{r[i]}21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                ws.Cells[$"{r[i]}25:{r[i]}32"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                ws.Cells[$"D{25 + i}:F{25 + i}"].Merge = true;
                ws.Cells[$"G{25 + i}:I{25 + i}"].Merge = true;

                ws.Cells[5 + i, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[5 + i, 13].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                ws.Cells[$"J{5 + i}:L{5 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"J{5 + i}:L{5 + i}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws.Cells[$"J{16 + i}:L{16 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"J{16 + i}:L{16 + i}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));

            }
            ws.Cells["M3:M10"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells["L14:L21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells["I25:I32"].Style.Border.Right.Style = ExcelBorderStyle.Thick;


            ws.Cells["C7:C10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C7:C10"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws.Cells["C18:C21"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C18:C21"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws.Cells["C29:C32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C29:C32"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));

            ws.Cells["C6:C9"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["C17:C20"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["C28:C31"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;



            int[] cH = { 3, 4, 14, 15 };

            for (int i = cH.GetLowerBound(0); i <= cH.GetUpperBound(0); i++)
            {
                ws.Cells[$"D{cH[i]}:F{cH[i]}"].Merge = true;
                ws.Cells[$"G{cH[i]}:I{cH[i]}"].Merge = true;
                ws.Cells[$"J{cH[i]}:L{cH[i]}"].Merge = true;
                ws.Cells[$"D{3 + i}:M{3 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"D{14 + i}:L{14 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"D{25 + i}:I{25 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            int[] cHH = { 5, 6, 16, 17, 27, 28 };

            for (int i = cHH.GetLowerBound(0); i <= cHH.GetUpperBound(0); i++)
            {
                ws.Cells[$"D{cHH[i]}:F{cHH[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"D{cHH[i]}:F{cHH[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws.Cells[$"G{cHH[i]}:I{cHH[i]}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"G{cHH[i]}:I{cHH[i]}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));

                ws.Cells[27, 4 + i].Value = lstSecond[i];
                ws.Cells[27, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[28, 4 + i].Value = "%";
                ws.Cells[28, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            }

            string[] H = { "D3", "G3", "J3", "M3", "D14", "G14", "J14", "G25", "D25" };
            string[] SH = { "D4", "G4", "J4", "M4", "D15", "G15", "J15", "G26", "D26" };
            string[] C = { "D", "E", "F", "G", "H", "I", "J", "K", "L" };
            for (int i = C.GetLowerBound(0); i <= C.GetUpperBound(0); i++)
            {
                ws.Cells[$"{C[i]}3:{C[i]}10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            for (int i = H.GetLowerBound(0); i <= H.GetUpperBound(0); i++)
            {
                ws.Cells[16, 4 + i].Value = lstSecond[i];
                ws.Cells[17, 4 + i].Value = "%";
                ws.Cells[16, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[17, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i % 2 != 0)
                {
                    ws.Cells[H[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[H[i]].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));
                    ws.Cells[SH[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[SH[i]].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));

                }
                else if (i % 2 == 0)
                {
                    ws.Cells[H[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[H[i]].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
                    ws.Cells[SH[i]].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[SH[i]].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));

                }
            }

            string[] Co1 = { "D3", "G3", "J3", "M3" };
            string[] Co2 = { "D4", "G4", "J4", "M4" };

            for (int i = Co1.GetLowerBound(0); i <= Co1.GetUpperBound(0); i++)
            {
                ws.Cells[$"{Co1[i]}"].Value = lstHeader[i];
                ws.Cells[$"{Co1[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co1[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"{Co2[i]}"].Value = "Quarter";
                ws.Cells[$"{Co2[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co2[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < 10; i++)
            {
                ws.Cells[5, 4 + i].Value = lstSecond[i];
                ws.Cells[5, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[6, 4 + i].Value = "%";
                ws.Cells[6, 4 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(4 + i).Width = 11;

            }


            string[] Co3 = { "D14", "G14", "J14" };
            string[] Co4 = { "D15", "G15", "J15" };
            for (int i = Co3.GetLowerBound(0); i <= Co3.GetUpperBound(0); i++)
            {
                ws.Cells[$"{Co3[i]}"].Value = lstHeader[i];
                ws.Cells[$"{Co3[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co3[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"{Co4[i]}"].Value = "Month/Year";
                ws.Cells[$"{Co4[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co4[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            string[] Co5 = { "D25", "G25" };
            string[] Co6 = { "D26", "G26" };
            for (int i = Co5.GetLowerBound(0); i <= Co5.GetUpperBound(0); i++)
            {
                ws.Cells[$"{Co5[i]}"].Value = lstHeader[i];
                ws.Cells[$"{Co5[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co5[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"{Co6[i]}"].Value = "Month/Year";
                ws.Cells[$"{Co6[i]}"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws.Cells[$"{Co6[i]}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "ReconciliationFactorsTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                FileStream fs = File.OpenWrite(dialog.FileName);
                fs.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                }
                IsEnabled1 = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Upload Error");
            }

        }

        public class F0
        {
            public string Quarter { get; set; }
            public double MillOre { get; set; }
            public double OLOre { get; set; }
            public double SLOre { get; set; }
            public double MillCuT { get; set; }
            public double OLCuT { get; set; }
            public double SLCuT { get; set; }
            public double MillCuFines { get; set; }
            public double OLCuFines { get; set; }
            public double SLCuFines { get; set; }
        }

        readonly List<F0> lstF0 = new List<F0>();

        public class F1
        {
            public string Quarter { get; set; }
            public double MillOre { get; set; }
            public double OLOre { get; set; }
            public double SLOre { get; set; }
            public double MillCuT { get; set; }
            public double OLCuT { get; set; }
            public double SLCuT { get; set; }
            public double MillCuFines { get; set; }
            public double OLCuFines { get; set; }
            public double SLCuFines { get; set; }
        }

        readonly List<F1> lstF1 = new List<F1>();

        public class F2
        {
            public string Quarter { get; set; }
            public double MillOre { get; set; }
            public double OLOre { get; set; }
            public double MillCuT { get; set; }
            public double OLCuT { get; set; }
            public double MillCuFines { get; set; }
            public double OLCuFines { get; set; }
        }

        readonly List<F2> lstF2 = new List<F2>();

        public class F3
        {
            public string Quarter { get; set; }
            public double MillCuFines { get; set; }
        }

        readonly List<F3> lstF3 = new List<F3>();

        private void LoadTemplate()
        {
            lstF0.Clear();
            lstF1.Clear();
            lstF2.Clear();
            lstF3.Clear();

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

                    ExcelWorksheet ws = pck.Workbook.Worksheets["Rolling Twelve Months"];

                    for (int i = 0; i < 4; i++)
                    {
                        lstF0.Add(new F0()
                        {
                            Quarter = ws.Cells[7 + i, 3].Value.ToString(),
                            MillOre = double.Parse(ws.Cells[7 + i, 4].Value.ToString()),
                            OLOre = double.Parse(ws.Cells[18 + i, 4].Value.ToString()),
                            SLOre = double.Parse(ws.Cells[29 + i, 4].Value.ToString()),
                            MillCuT = double.Parse(ws.Cells[7 + i, 5].Value.ToString()),
                            OLCuT = double.Parse(ws.Cells[18 + i, 5].Value.ToString()),
                            SLCuT = double.Parse(ws.Cells[29 + i, 5].Value.ToString()),
                            MillCuFines = double.Parse(ws.Cells[7 + i, 6].Value.ToString()),
                            OLCuFines = double.Parse(ws.Cells[18 + i, 6].Value.ToString()),
                            SLCuFines = double.Parse(ws.Cells[29 + i, 6].Value.ToString())
                        });
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        lstF1.Add(new F1()
                        {
                            Quarter = ws.Cells[7 + i, 3].Value.ToString(),
                            MillOre = double.Parse(ws.Cells[7 + i, 7].Value.ToString()),
                            OLOre = double.Parse(ws.Cells[18 + i, 7].Value.ToString()),
                            SLOre = double.Parse(ws.Cells[29 + i, 7].Value.ToString()),
                            MillCuT = double.Parse(ws.Cells[7 + i, 8].Value.ToString()),
                            OLCuT = double.Parse(ws.Cells[18 + i, 8].Value.ToString()),
                            SLCuT = double.Parse(ws.Cells[29 + i, 8].Value.ToString()),
                            MillCuFines = double.Parse(ws.Cells[7 + i, 9].Value.ToString()),
                            OLCuFines = double.Parse(ws.Cells[18 + i, 9].Value.ToString()),
                            SLCuFines = double.Parse(ws.Cells[29 + i, 9].Value.ToString())
                        });
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        lstF2.Add(new F2()
                        {
                            Quarter = ws.Cells[7 + i, 3].Value.ToString(),
                            MillOre = double.Parse(ws.Cells[7 + i, 10].Value.ToString()),
                            OLOre = double.Parse(ws.Cells[18 + i, 10].Value.ToString()),
                            MillCuT = double.Parse(ws.Cells[7 + i, 11].Value.ToString()),
                            OLCuT = double.Parse(ws.Cells[18 + i, 11].Value.ToString()),
                            MillCuFines = double.Parse(ws.Cells[7 + i, 12].Value.ToString()),
                            OLCuFines = double.Parse(ws.Cells[18 + i, 12].Value.ToString())
                        });
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        lstF3.Add(new F3()
                        {
                            Quarter = ws.Cells[7 + i, 3].Value.ToString(),
                            MillCuFines = double.Parse(ws.Cells[7 + i, 13].Value.ToString())
                        });
                    }

                    pck.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\RollingTwelveMonthData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {                   

                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        ExcelWorksheet ws2 = pck2.Workbook.Worksheets["F0"];
                        DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00).AddMilliseconds(000);
                        int lastRow1 = ws2.Dimension.End.Row + 1;

                        for (int i = 0; i < lstF0.Count; i++)
                        {
                            ws2.Cells[i + lastRow1, 1].Value = newDate;
                            ws2.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            //ws2.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss.000";
                            ws2.Cells[i + lastRow1, 2].Value = lstF0[i].Quarter;
                            ws2.Cells[i + lastRow1, 3].Value = lstF0[i].MillOre;
                            ws2.Cells[i + lastRow1, 4].Value = lstF0[i].OLOre;
                            ws2.Cells[i + lastRow1, 5].Value = lstF0[i].SLOre;
                            ws2.Cells[i + lastRow1, 6].Value = lstF0[i].MillCuT;
                            ws2.Cells[i + lastRow1, 7].Value = lstF0[i].OLCuT;
                            ws2.Cells[i + lastRow1, 8].Value = lstF0[i].SLCuT;
                            ws2.Cells[i + lastRow1, 9].Value = lstF0[i].MillCuFines;
                            ws2.Cells[i + lastRow1, 10].Value = lstF0[i].OLCuFines;
                            ws2.Cells[i + lastRow1, 11].Value = lstF0[i].SLCuFines;
                        }

                        ExcelWorksheet ws3 = pck2.Workbook.Worksheets["F1"];
                        int lastRow2 = ws3.Dimension.End.Row + 1;

                        for (int i = 0; i < lstF1.Count; i++)
                        {
                            ws3.Cells[i + lastRow2, 1].Value = newDate;
                            ws3.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            // ws3.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss.000";
                            ws3.Cells[i + lastRow2, 2].Value = lstF1[i].Quarter;
                            ws3.Cells[i + lastRow2, 3].Value = lstF1[i].MillOre;
                            ws3.Cells[i + lastRow2, 4].Value = lstF1[i].OLOre;
                            ws3.Cells[i + lastRow2, 5].Value = lstF1[i].SLOre;
                            ws3.Cells[i + lastRow2, 6].Value = lstF1[i].MillCuT;
                            ws3.Cells[i + lastRow2, 7].Value = lstF1[i].OLCuT;
                            ws3.Cells[i + lastRow2, 8].Value = lstF1[i].SLCuT;
                            ws3.Cells[i + lastRow2, 9].Value = lstF1[i].MillCuFines;
                            ws3.Cells[i + lastRow2, 10].Value = lstF1[i].OLCuFines;
                            ws3.Cells[i + lastRow2, 11].Value = lstF1[i].SLCuFines;
                        }

                        ExcelWorksheet ws4 = pck2.Workbook.Worksheets["F2"];
                        int lastRow3 = ws4.Dimension.End.Row + 1;

                        for (int i = 0; i < lstF2.Count; i++)
                        {
                            ws4.Cells[i + lastRow3, 1].Value = newDate;
                            ws4.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            //ws4.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss.000";
                            ws4.Cells[i + lastRow3, 2].Value = lstF2[i].Quarter;
                            ws4.Cells[i + lastRow3, 3].Value = lstF2[i].MillOre;
                            ws4.Cells[i + lastRow3, 4].Value = lstF2[i].OLOre;
                            ws4.Cells[i + lastRow3, 5].Value = lstF2[i].MillCuT;
                            ws4.Cells[i + lastRow3, 6].Value = lstF2[i].OLCuT;
                            ws4.Cells[i + lastRow3, 7].Value = lstF2[i].MillCuFines;
                            ws4.Cells[i + lastRow3, 8].Value = lstF2[i].OLCuFines;
                        }

                        ExcelWorksheet ws5 = pck2.Workbook.Worksheets["F3"];
                        int lastRow4 = ws5.Dimension.End.Row + 1;

                        for (int i = 0; i < lstF3.Count; i++)
                        {
                            ws5.Cells[i + lastRow4, 1].Value = newDate;
                            ws5.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            //ws5.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss.000";
                            ws5.Cells[i + lastRow4, 2].Value = lstF3[i].Quarter;
                            ws5.Cells[i + lastRow4, 3].Value = lstF3[i].MillCuFines;
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

    }
}
