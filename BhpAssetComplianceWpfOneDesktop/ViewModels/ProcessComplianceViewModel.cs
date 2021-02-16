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
using System.Globalization;
using OfficeOpenXml.Style;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class ProcessComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.ProcessCompliance;

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

        public ProcessComplianceViewModel()
        {
            Date = DateTime.Now;
            IsEnabled1 = false;
            GenerarT = new DelegateCommand(GenerateTemplate);
            CargarT = new DelegateCommand(LoadTemplate).ObservesCanExecute(() => IsEnabled1);
        }

        private void GenerateTemplate()
        {
            List<string> lstPhase = new List<string>() { "Phase", "Ore to Mill Budget (Mt)", "Ore to Mill Actual (Mt)", "Hardness Budget (min)", "Hardness Actual (min)" };
            List<string> lstPhase2 = new List<string>() { "Phase", "Recovery Budget (%)", "Recovery Actual (%)", "Feed Cu Budget (%)", "Feed Cu Actual (%)" };
            List<string> lstFeed = new List<string>() { "Feed Grade", "Stacked Ore (kt)", "CuT (%)", "Cathodes (t)", "Distribution", "Expit", "Average CuT", "Stocks" };
            List<string> lstDist = new List<string>() { "Budget", "Actual", "Compliance %", "Budget %", "Actual %" };

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Process Compliance Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("Ore to Mill");
            ws.Cells["B1:C1"].Style.Font.Bold = true;
            ws.Cells["B1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B1:C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
            ws.Cells["B1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B1"].Value = "Budget (min)";
            ws.Cells["C1"].Value = "Actual (min)";

            for (int i = 0; i < 3; i++)
            {
                ws.Cells[1, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[2, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
            ws.Cells["A2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A2"].Value = "SPI Global";

            ws.Cells["A5:E5"].Style.Font.Bold = true;
            ws.Cells["A5:E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
            ws.Cells["B5:E5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
            ws.Cells["B5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < 17; i++)
            {
                ws.Cells[$"A{4 + i}:E{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] D = { "A", "B", "C", "D", "E" };
            for (int i = D.GetLowerBound(0); i <= D.GetUpperBound(0); i++)
            {
                ws.Cells[5, 1 + i].Value = lstPhase[i];
                ws.Cells[$"{D[i]}5:{D[i]}20"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Column(1 + i).Width = 19;
            }
            ws.Column(1).Width = 11;

            var ws4 = pck.Workbook.Worksheets.Add("Recovery");
            ws4.Cells["B1:C1"].Style.Font.Bold = true;
            ws4.Cells["B1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws4.Cells["B1:C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
            ws4.Cells["B1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws4.Cells["B1"].Value = "Budget (%)";
            ws4.Cells["C1"].Value = "Actual (%)";

            for (int i = 0; i < 3; i++)
            {
                ws4.Cells[1, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws4.Cells[2, 1 + i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws4.Cells[1, 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            ws4.Cells["A2"].Style.Font.Bold = true;
            ws4.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws4.Cells["A2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
            ws4.Cells["A2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws4.Cells["A2"].Value = "Rec Global";

            ws4.Cells["A5:E5"].Style.Font.Bold = true;
            ws4.Cells["A5:E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws4.Cells["A5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
            ws4.Cells["B5:E5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
            ws4.Cells["B5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < 17; i++)
            {
                ws4.Cells[$"A{4 + i}:E{4 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            string[] G = { "A", "B", "C", "D", "E" };
            for (int i = G.GetLowerBound(0); i <= G.GetUpperBound(0); i++)
            {
                ws4.Cells[5, 1 + i].Value = lstPhase2[i];
                ws4.Cells[$"{G[i]}5:{G[i]}20"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws4.Column(1 + i).Width = 19;
            }
            ws4.Column(1).Width = 11;

            var ws2 = pck.Workbook.Worksheets.Add("OLAP");
            ws2.Cells["A1:H1"].Style.Font.Bold = true;
            ws2.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["F1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["A2:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Cells["F2:F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws2.Column(1).Width = 14;
            ws2.Column(6).Width = 14;
            string[] E = { "A", "B", "C", "D", "F", "G", "H" };

            for (int i = 0; i < 4; i++)
            {
                ws2.Cells[1 + i, 1].Value = lstFeed[i];
                ws2.Cells[1 + i, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
                ws2.Cells[1 + i, 6].Value = lstFeed[4 + i];
                ws2.Cells[1 + i, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
                ws2.Column(2 + i).Width = 12;
                ws2.Column(7 + i).Width = 12;
                ws2.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"F{1 + i}:H{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"{E[i]}1:{E[i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            ws2.Cells["G1:H1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));

            for (int i = 0; i < 3; i++)
            {
                ws2.Cells[1, 2 + i].Value = lstDist[i];
                ws2.Cells[1, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
                ws2.Cells[1, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells[$"{E[4 + i]}1:{E[4 + i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            ws2.Cells[1, 7].Value = lstDist[3];
            ws2.Cells[1, 8].Value = lstDist[4];
            ws2.Cells["G1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var ws3 = pck.Workbook.Worksheets.Add("Sulphide");
            ws3.Cells["A1:H1"].Style.Font.Bold = true;
            ws3.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["F1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["A2:A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Cells["F2:F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws3.Column(1).Width = 14;
            ws3.Column(6).Width = 14;
            string[] F = { "A", "B", "C", "D", "F", "G", "H" };

            for (int i = 0; i < 4; i++)
            {
                ws3.Cells[1 + i, 1].Value = lstFeed[i];
                ws3.Cells[1 + i, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
                ws3.Cells[1 + i, 6].Value = lstFeed[4 + i];
                ws3.Cells[1 + i, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFA153"));
                ws3.Column(2 + i).Width = 12;
                ws3.Column(7 + i).Width = 12;
                ws3.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws3.Cells[$"F{1 + i}:H{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws3.Cells[$"{F[i]}1:{F[i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            ws3.Cells["G1:H1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));

            for (int i = 0; i < 3; i++)
            {
                ws3.Cells[1, 2 + i].Value = lstDist[i];
                ws3.Cells[1, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#9BB3C1"));
                ws3.Cells[1, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells[$"{F[4 + i]}1:{F[4 + i]}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            ws3.Cells[1, 7].Value = lstDist[3];
            ws3.Cells[1, 8].Value = lstDist[4];
            ws3.Cells["G1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "ProcessComplianceTemplate.xlsx",
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

        public class OretoMill
        {
            public double SpiGlobalBudget { get; set; }
            public double SpiGlobalActual { get; set; }
            public string Phase { get; set; }
            public double OretoMillBudget { get; set; }
            public double OretoMillActual { get; set; }
            public double HardnessBudget { get; set; }
            public double HardnessActual { get; set; }
        }

        readonly List<OretoMill> lstOretoMill = new List<OretoMill>();

        public class Recovery
        {
            public double RecGlobalBudget { get; set; }
            public double RecGlobalActual { get; set; }
            public string Phase { get; set; }
            public double RecoveryBudget { get; set; }
            public double RecoveryActual { get; set; }
            public double FeedCuBudget { get; set; }
            public double FeedCuActual { get; set; }
        }

        readonly List<Recovery> lstRecovery = new List<Recovery>();

        public class OLAP
        {
            public string FeedGrade { get; set; }
            public double Budget { get; set; }
            public double Actual { get; set; }
            public double Compliance { get; set; }
            public string Distribution { get; set; }
            public double DistributionBudget { get; set; }
            public double DistributionActual { get; set; }
        }

        readonly List<OLAP> lstOLAP = new List<OLAP>();

        public class Sulphide
        {
            public string FeedGrade { get; set; }
            public double Budget { get; set; }
            public double Actual { get; set; }
            public double Compliance { get; set; }
            public string Distribution { get; set; }
            public double DistributionBudget { get; set; }
            public double DistributionActual { get; set; }
        }

        readonly List<Sulphide> lstSulphide = new List<Sulphide>();

        private void LoadTemplate()
        {
            lstOretoMill.Clear();
            lstRecovery.Clear();
            lstOLAP.Clear();
            lstSulphide.Clear();

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

                    ExcelWorksheet ws = pck.Workbook.Worksheets["Ore to Mill"];
                    int rows = ws.Dimension.Rows;

                    for (int i = 0; i < rows; i++)
                    {
                        if (ws.Cells[i + 6, 1].Value != null)
                        {
                            if (ws.Cells[2, 2].Value == null)
                            {
                                ws.Cells[2, 2].Value = -99;
                            }
                            if (ws.Cells[2, 3].Value == null)
                            {
                                ws.Cells[2, 3].Value = -99;
                            }

                            for (int j = 0; j < 4; j++)
                            {
                                if (ws.Cells[6 + i, 2 + j].Value == null)
                                {
                                    ws.Cells[6 + i, 2 + j].Value = -99;
                                }
                            }

                            lstOretoMill.Add(new OretoMill()
                            {
                                SpiGlobalBudget = double.Parse(ws.Cells[2, 2].Value.ToString()),
                                SpiGlobalActual = double.Parse(ws.Cells[2, 3].Value.ToString()),
                                Phase = ws.Cells[6 + i, 1].Value.ToString(),
                                OretoMillBudget = double.Parse(ws.Cells[6 + i, 2].Value.ToString()),
                                OretoMillActual = double.Parse(ws.Cells[6 + i, 3].Value.ToString()),
                                HardnessBudget = double.Parse(ws.Cells[6 + i, 4].Value.ToString()),
                                HardnessActual = double.Parse(ws.Cells[6 + i, 5].Value.ToString())
                            });
                        }
                    }

                    ExcelWorksheet ws4 = pck.Workbook.Worksheets["Recovery"];
                    int rows2 = ws4.Dimension.Rows;

                    for (int i = 0; i < rows2; i++)
                    {
                        if (ws4.Cells[i + 6, 1].Value != null)
                        {
                            if (ws4.Cells[2, 2].Value == null)
                            {
                                ws4.Cells[2, 2].Value = -99;
                            }

                            if (ws4.Cells[2, 3].Value == null)
                            {
                                ws4.Cells[2, 3].Value = -99;
                            }

                            for (int j = 0; j < 4; j++)
                            {
                                if (ws4.Cells[6 + i, 2 + j].Value == null)
                                {
                                    ws4.Cells[6 + i, 2 + j].Value = -99;
                                }
                            }

                            lstRecovery.Add(new Recovery()
                            {
                                RecGlobalBudget = double.Parse(ws4.Cells[2, 2].Value.ToString()),
                                RecGlobalActual = double.Parse(ws4.Cells[2, 3].Value.ToString()),
                                Phase = ws4.Cells[6 + i, 1].Value.ToString(),
                                RecoveryBudget = double.Parse(ws4.Cells[6 + i, 2].Value.ToString()) / 100,
                                RecoveryActual = double.Parse(ws4.Cells[6 + i, 3].Value.ToString()) / 100,
                                FeedCuBudget = double.Parse(ws4.Cells[6 + i, 4].Value.ToString()) / 100,
                                FeedCuActual = double.Parse(ws4.Cells[6 + i, 5].Value.ToString()) / 100
                            });
                        }
                    }

                    ExcelWorksheet ws2 = pck.Workbook.Worksheets["OLAP"];
                    for (int i = 0; i < 3; i++)
                    {

                        for (int j = 0; j < 3; j++)
                        {
                            if (ws2.Cells[2 + i, 2 + j].Value == null)
                            {
                                ws2.Cells[2 + i, 2 + j].Value = -99;
                            }
                        }

                        for (int j = 0; j < 2; j++)
                        {
                            if (ws2.Cells[2 + i, 7 + j].Value == null)
                            {
                                ws2.Cells[2 + i, 7 + j].Value = -99;
                            }
                        }

                        lstOLAP.Add(new OLAP()
                        {
                            FeedGrade = ws2.Cells[2 + i, 1].Value.ToString(),
                            Budget = double.Parse(ws2.Cells[2 + i, 2].Value.ToString()),
                            Actual = double.Parse(ws2.Cells[2 + i, 3].Value.ToString()),
                            Compliance = double.Parse(ws2.Cells[2 + i, 4].Value.ToString()),
                            Distribution = ws2.Cells[2 + i, 6].Value.ToString(),
                            DistributionBudget = double.Parse(ws2.Cells[2 + i, 7].Value.ToString()) / 100,
                            DistributionActual = double.Parse(ws2.Cells[2 + i, 8].Value.ToString()) / 100
                        });
                    }

                    ExcelWorksheet ws3 = pck.Workbook.Worksheets["Sulphide"];
                    for (int i = 0; i < 3; i++)
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            if (ws3.Cells[2 + i, 2 + j].Value == null)
                            {
                                ws3.Cells[2 + i, 2 + j].Value = -99;
                            }
                        }

                        for (int j = 0; j < 2; j++)
                        {
                            if (ws3.Cells[2 + i, 7 + j].Value == null)
                            {
                                ws3.Cells[2 + i, 7 + j].Value = -99;
                            }
                        }

                        lstSulphide.Add(new Sulphide()
                        {
                            FeedGrade = ws3.Cells[2 + i, 1].Value.ToString(),
                            Budget = double.Parse(ws3.Cells[2 + i, 2].Value.ToString()),
                            Actual = double.Parse(ws3.Cells[2 + i, 3].Value.ToString()),
                            Compliance = double.Parse(ws3.Cells[2 + i, 4].Value.ToString()),
                            Distribution = ws3.Cells[2 + i, 6].Value.ToString(),
                            DistributionBudget = double.Parse(ws3.Cells[2 + i, 7].Value.ToString()) / 100,
                            DistributionActual = double.Parse(ws3.Cells[2 + i, 8].Value.ToString()) / 100
                        });
                    }

                    pck.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\ProcessComplianceData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        ExcelWorksheet w2s = pck2.Workbook.Worksheets["Ore to Mill"];
                        DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00).AddMilliseconds(000);
                        int lastRow1 = w2s.Dimension.End.Row + 1;

                        for (int i = 0; i < lstOretoMill.Count; i++)
                        {
                            w2s.Cells[i + lastRow1, 1].Value = newDate;
                            w2s.Cells[i + lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            w2s.Cells[i + lastRow1, 2].Value = lstOretoMill[i].SpiGlobalBudget;
                            w2s.Cells[i + lastRow1, 3].Value = lstOretoMill[i].SpiGlobalActual;
                            w2s.Cells[i + lastRow1, 4].Value = lstOretoMill[i].Phase;
                            w2s.Cells[i + lastRow1, 5].Value = lstOretoMill[i].OretoMillBudget;
                            w2s.Cells[i + lastRow1, 6].Value = lstOretoMill[i].OretoMillActual;
                            w2s.Cells[i + lastRow1, 7].Value = lstOretoMill[i].HardnessBudget;
                            w2s.Cells[i + lastRow1, 8].Value = lstOretoMill[i].HardnessActual;
                        }

                        ExcelWorksheet w2s4 = pck2.Workbook.Worksheets["Recovery"];
                        int lastRow2 = w2s4.Dimension.End.Row + 1;

                        for (int i = 0; i < lstRecovery.Count; i++)
                        {
                            w2s4.Cells[i + lastRow2, 1].Value = newDate;
                            w2s4.Cells[i + lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            w2s4.Cells[i + lastRow2, 2].Value = lstRecovery[i].RecGlobalBudget;
                            w2s4.Cells[i + lastRow2, 3].Value = lstRecovery[i].RecGlobalActual;
                            w2s4.Cells[i + lastRow2, 4].Value = lstRecovery[i].Phase;
                            w2s4.Cells[i + lastRow2, 5].Value = lstRecovery[i].RecoveryBudget;
                            w2s4.Cells[i + lastRow2, 6].Value = lstRecovery[i].RecoveryActual;
                            w2s4.Cells[i + lastRow2, 7].Value = lstRecovery[i].FeedCuBudget;
                            w2s4.Cells[i + lastRow2, 8].Value = lstRecovery[i].FeedCuActual;
                        }

                        ExcelWorksheet w2s2 = pck2.Workbook.Worksheets["OLAP"];
                        int lastRow3 = w2s2.Dimension.End.Row + 1;

                        for (int i = 0; i < lstOLAP.Count; i++)
                        {
                            w2s2.Cells[i + lastRow3, 1].Value = newDate;
                            w2s2.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            w2s2.Cells[i + lastRow3, 2].Value = lstOLAP[i].FeedGrade;
                            w2s2.Cells[i + lastRow3, 3].Value = lstOLAP[i].Budget;
                            w2s2.Cells[i + lastRow3, 4].Value = lstOLAP[i].Actual;
                            w2s2.Cells[i + lastRow3, 5].Value = lstOLAP[i].Compliance;
                            w2s2.Cells[i + lastRow3, 6].Value = lstOLAP[i].Distribution;
                            w2s2.Cells[i + lastRow3, 7].Value = lstOLAP[i].DistributionBudget;
                            w2s2.Cells[i + lastRow3, 8].Value = lstOLAP[i].DistributionActual;
                        }

                        ExcelWorksheet w2s3 = pck2.Workbook.Worksheets["Sulphide"];
                        int lastRow4 = w2s3.Dimension.End.Row + 1;

                        for (int i = 0; i < lstSulphide.Count; i++)
                        {
                            w2s3.Cells[i + lastRow4, 1].Value = newDate;
                            w2s3.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            w2s3.Cells[i + lastRow4, 2].Value = lstSulphide[i].FeedGrade;
                            w2s3.Cells[i + lastRow4, 3].Value = lstSulphide[i].Budget;
                            w2s3.Cells[i + lastRow4, 4].Value = lstSulphide[i].Actual;
                            w2s3.Cells[i + lastRow4, 5].Value = lstSulphide[i].Compliance;
                            w2s3.Cells[i + lastRow4, 6].Value = lstSulphide[i].Distribution;
                            w2s3.Cells[i + lastRow4, 7].Value = lstSulphide[i].DistributionBudget;
                            w2s3.Cells[i + lastRow4, 8].Value = lstSulphide[i].DistributionActual;
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
