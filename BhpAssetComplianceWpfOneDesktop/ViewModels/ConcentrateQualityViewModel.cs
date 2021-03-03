using BhpAssetComplianceWpfOneDesktop.Constants;
using BhpAssetComplianceWpfOneDesktop.Models.ConcentrateQualityModels;
using BhpAssetComplianceWpfOneDesktop.Resources;
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
    public class ConcentrateQualityViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.ConcentrateQuality;
        protected override string MyPosterIcon { get; set; } = IconKeys.ConcentrateQuality;

        // TODO: Borrar variables que no se esten usando y usarlas como se presenta en el sistema
        public string generateContent { get; set; } = StringResources.GenerateTemplate;
        public string loadContent { get; set; } = StringResources.LoadTemplate;
        public string dateContent { get; set; } = StringResources.Date;
        public string fiscalYearContent { get; set; } = StringResources.FiscalYear;
        public string actualContent { get; set; } = StringResources.ActualMonthFreightTemplate;
        public string budgetContent { get; set; } = StringResources.BudgetFreightTemplate;

        // TODO: Colocar nombres que hagan sentido
        private string _UpdateA; // TODO: Respetar en la variables privadas la notación _updateA
        public string UpdateA
        {
            get { return _UpdateA; }
            set { SetProperty(ref _UpdateA, value); }
        }

        private string _UpdateB;
        // TODO: Poner nombre más autoexplicativos
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
        // TODO: Nombres de variables autoexplicativos
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

        // TODO: Idioma y evitar abreviaturas 
        public DelegateCommand GenerarAFT { get; set; }
        public DelegateCommand CargarAFT { get; set; }
        public DelegateCommand GenerarBFT { get; set; }
        public DelegateCommand CargarBFT { get; set; }

        public ConcentrateQualityViewModel()
        {
            Date = DateTime.Now;
            FiscalYear = Date.Year;
            IsEnabled1 = false;
            IsEnabled2 = false;
            GenerarAFT = new DelegateCommand(GenerateActualFreightTemplate);
            CargarAFT = new DelegateCommand(LoadActualFreightTemplate).ObservesCanExecute(() => IsEnabled1);
            GenerarBFT = new DelegateCommand(GenerateBudgetFreightTemplate);
            CargarBFT = new DelegateCommand(LoadBudgetFreightTemplate).ObservesCanExecute(() => IsEnabled2);
        }

        private void GenerateActualFreightTemplate()
        {
            // TODO: Cambiar a headers y usar var
            var lstHeader = new List<string>() { "Nombre M/N", "N°", "Inicio embarque", "Termino embarque" };
            List<string> lstItem = new List<string>() { "WMT", "DMT", "Moisture.", "Cu", "As", "Fe", "Au", "Ag", "S", "Insol.", "Cd", "Zn", "Hg", "SiO2", "Al2O3", "Sb", "Mo" };
            List<string> lstUnit = new List<string>() { "Pesometer t", "Pesometer t", "%", "%", "%", "%", "g/t", "g/t", "%", "%", "%", "%", "g/t", "%", "%", "%", "%" };

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Real Month Freight Template";
            pck.Workbook.Properties.Company = "BHP";

            // TODO: usar mejores variables de números
            var ws = pck.Workbook.Worksheets.Add("RealFreight"); // TODO: Constantes
            ws.Cells["A2:U2"].Style.Font.Bold = true; // TODO: Evitar usar referencias "A1" y usar números
            ws.Cells["A2:U2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));  // TODO: Utilizar constantes para colores
            ws.Cells["E3:U3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells["A2:U2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["E3:U3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Column(1).Width = 22;

            ws.Cells["A2:A3"].Merge = true;
            ws.Cells["B2:B3"].Merge = true;
            ws.Cells["C2:C3"].Merge = true;
            ws.Cells["D2:D3"].Merge = true;

            for (int i = 0; i < lstHeader.Count; i++)
            {
                ws.Cells[2, i + 1].Value = lstHeader[i];
                ws.Cells[2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[2, i + 1].Style.WrapText = true;
            }

            for (int i = 0; i < lstItem.Count; i++)
            {
                ws.Cells[2, i + 5].Value = lstItem[i];
                ws.Cells[3, i + 5].Value = lstUnit[i];
            }

            for (int i = 1; i < 22; i++)
            {
                if (i % 2 != 0)
                {
                    ws.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
                }
                else if (i % 2 == 0)
                {
                    ws.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));
                }
                ws.Column(1 + i).Width = 16;
            }

            for (int i = 1; i < 18; i++)
            {
                if (i % 2 != 0)
                {
                    ws.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                }
                else if (i % 2 == 0)
                {
                    ws.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));
                }
            }

            ws.Cells["A2:U2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < 14; i++)
            {
                ws.Cells[i + 2, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "FreightRealTemplate.xlsx",
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


        // TODO: Mover este campo a arriba y usar la notación _freights
        readonly List<ConcentrateQualityFreight> lstFreights = new List<ConcentrateQualityFreight>();

        private void LoadActualFreightTemplate()
        {
            lstFreights.Clear();
            // TODO: Usar var
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
                    ExcelWorksheet ws = pck.Workbook.Worksheets["RealFreight"];

                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    int rows = ws.Dimension.Rows;

                    for (int i = 1; i < rows; i++)
                    {
                        if (ws.Cells[i + 3, 1].Value != null)
                        {

                            for (int j = 0; j < 17; j++)
                            {
                                if (ws.Cells[3 + i, 5 + j].Value == null)
                                {
                                    ws.Cells[3 + i, 5 + j].Value = -99;
                                }

                            }

                            lstFreights.Add(new ConcentrateQualityFreight()
                            {

                                Name = ws.Cells[3 + i, 1].Value.ToString(),
                                Number = Int32.Parse(ws.Cells[3 + i, 2].Value.ToString()),
                                Start = Convert.ToDateTime(ws.Cells[3 + i, 3].Value.ToString()),
                                End = Convert.ToDateTime(ws.Cells[3 + i, 4].Value.ToString()),
                                WMT = double.Parse(ws.Cells[3 + i, 5].Value.ToString()),
                                DMT = double.Parse(ws.Cells[3 + i, 6].Value.ToString()),
                                Moisture = double.Parse(ws.Cells[3 + i, 7].Value.ToString()) / 100,
                                Cu = double.Parse(ws.Cells[3 + i, 8].Value.ToString()) / 100,
                                As = double.Parse(ws.Cells[3 + i, 9].Value.ToString()) * 10000,
                                Fe = double.Parse(ws.Cells[3 + i, 10].Value.ToString()) / 100,
                                Au = double.Parse(ws.Cells[3 + i, 11].Value.ToString()),
                                Ag = double.Parse(ws.Cells[3 + i, 12].Value.ToString()),
                                S = double.Parse(ws.Cells[3 + i, 13].Value.ToString()) / 100,
                                Insoluble = double.Parse(ws.Cells[3 + i, 14].Value.ToString()) / 100,
                                Cd = double.Parse(ws.Cells[3 + i, 15].Value.ToString()) * 10000,
                                Zn = double.Parse(ws.Cells[3 + i, 16].Value.ToString()) * 10000,
                                Hg = double.Parse(ws.Cells[3 + i, 17].Value.ToString()),
                                SiO2 = double.Parse(ws.Cells[3 + i, 18].Value.ToString()) / 100,
                                Al2O3 = double.Parse(ws.Cells[3 + i, 19].Value.ToString()) / 100,
                                Sb = double.Parse(ws.Cells[3 + i, 20].Value.ToString()) * 10000,
                                Mo = double.Parse(ws.Cells[3 + i, 21].Value.ToString())
                            });

                        }

                    }

                    pck.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\FreightData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        ExcelWorksheet ws2 = pck2.Workbook.Worksheets["RealMonth"];

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        int lastRow = ws2.Dimension.End.Row + 1;
                        DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00).AddMilliseconds(000);

                        for (int i = 0; i < lstFreights.Count; i++)
                        {
                            ws2.Cells[i + lastRow, 1].Value = newDate;
                            ws2.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 2].Value = lstFreights[i].Name;
                            ws2.Cells[i + lastRow, 3].Value = lstFreights[i].Number;
                            ws2.Cells[i + lastRow, 4].Value = lstFreights[i].Start;
                            ws2.Cells[i + lastRow, 4].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 5].Value = lstFreights[i].End;
                            ws2.Cells[i + lastRow, 5].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 6].Value = lstFreights[i].WMT;
                            ws2.Cells[i + lastRow, 7].Value = lstFreights[i].DMT;
                            ws2.Cells[i + lastRow, 8].Value = lstFreights[i].Moisture;
                            ws2.Cells[i + lastRow, 9].Value = lstFreights[i].Cu;
                            ws2.Cells[i + lastRow, 10].Value = lstFreights[i].As;
                            ws2.Cells[i + lastRow, 11].Value = lstFreights[i].Fe;
                            ws2.Cells[i + lastRow, 12].Value = lstFreights[i].Au;
                            ws2.Cells[i + lastRow, 13].Value = lstFreights[i].Ag;
                            ws2.Cells[i + lastRow, 14].Value = lstFreights[i].S;
                            ws2.Cells[i + lastRow, 15].Value = lstFreights[i].Insoluble;
                            ws2.Cells[i + lastRow, 16].Value = lstFreights[i].Cd;
                            ws2.Cells[i + lastRow, 17].Value = lstFreights[i].Zn;
                            ws2.Cells[i + lastRow, 18].Value = lstFreights[i].Hg;
                            ws2.Cells[i + lastRow, 19].Value = lstFreights[i].SiO2;
                            ws2.Cells[i + lastRow, 20].Value = lstFreights[i].Al2O3;
                            ws2.Cells[i + lastRow, 21].Value = lstFreights[i].Sb;
                            ws2.Cells[i + lastRow, 22].Value = lstFreights[i].Mo;
                        }

                        byte[] fileText2 = pck2.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText2);

                        UpdateA = $"{StringResources.Updated}: {DateTime.Now}";

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Upload Error");
                    }
                }

            }

        }

        private void GenerateBudgetFreightTemplate()
        {
            List<string> lstHeader = new List<string>() { "Item", "Unit", "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            List<string> lstItem = new List<string>() { "Au", "Ag", "Mo", "As", "Cd", "Pb", "Zn", "Bi", "Sb", "Fe Conc", "Fe", "Py Conc", "Py", "S2", "Concentrate Grade" };
            List<string> lstUnit = new List<string>() { "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "ppm", "%", "%", "%", "%", "%", "%" };

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Budget Freight Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("BudgetFreight");

            ws.Cells["B2:O2"].Merge = true;
            ws.Cells["B2"].Value = $"FY{FiscalYear}";

            ws.Column(2).Style.Font.Bold = true;
            ws.Column(3).Style.Font.Bold = true;

            ws.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#AEAAAA"));
            ws.Cells["B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < lstHeader.Count; i++)
            {
                ws.Cells[3, i + 2].Value = lstHeader[i];
                ws.Cells[3, i + 2].Style.Font.Bold = true;
                ws.Cells[3, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[3, i + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E7E6E6"));
            }

            ws.Cells["B2:O2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Column(2).Width = 22;
            ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < 17; i++)
            {
                ws.Cells[i + 2, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"B{2 + i}:O{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"B{2 + i}:O{2 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Column(3 + i).Width = 11;
                ws.Cells[3, i + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < lstItem.Count; i++)
            {
                ws.Cells[4 + i, 2].Value = lstItem[i];
                ws.Cells[4 + i, 3].Value = lstUnit[i];
                ws.Cells[$"B{4 + i}:C{4 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"B{4 + i}:C{4 + i}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E7E6E6"));
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "FreightBudgetTemplate.xlsx",
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
        //TODO: Formatear esto con un nombre más comprensible y sacar de la clase y mover a la carpeta modelos
        public class Item
        {
            public DateTime Date { get; set; }
            public double Au { get; set; }
            public double Ag { get; set; }
            public double Mo { get; set; }
            public double As { get; set; }
            public double Cd { get; set; }
            public double Pb { get; set; }
            public double Zn { get; set; }
            public double Bi { get; set; }
            public double Sb { get; set; }
            public double FeConc { get; set; }
            public double Fe { get; set; }
            public double PyConc { get; set; }
            public double Py { get; set; }
            public double S2 { get; set; }
            public double ConcentrateGrade { get; set; }
        }

        // TODO: Refactoring para un nombres comprensibles y subir a la zona de propiedades
        readonly List<Item> lstItems = new List<Item>();
        private void LoadBudgetFreightTemplate()
        {
            lstItems.Clear();
            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (op.ShowDialog() == true)
            {
                try
                {
                    // TODO: Usar nomenclatura para variable filePath
                    FileInfo FilePath = new FileInfo(op.FileName);
                    ExcelPackage pck = new ExcelPackage(FilePath);
                    ExcelWorksheet ws = pck.Workbook.Worksheets["BudgetFreight"]; // TODO: Llevar a un archivo de constantes

                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    // TODO: Usar variable
                    DateTime Db = DateTime.Now;

                    for (int i = 0; i < 12; i++)
                    {
                        // TODO: Evitar variables con letras, usar variable _month
                        int M = DateTime.ParseExact(ws.Cells[3, 4 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;

                        // TODO: utilizar una función en utility para esta conversion
                        if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i == 5)
                        {
                            Db = new DateTime(FiscalYear - 1, M, 1, 00, 00, 00).AddMilliseconds(000); // TODO: No es necesario el add milliseconds
                        }
                        else if (i == 6 || i == 7 || i == 8 || i == 9 || i == 10 || i == 11)
                        {
                            Db = new DateTime(FiscalYear, M, 1, 00, 00, 00).AddMilliseconds(000);
                        }


                        for (int j = 0; j < 17; j++)
                        {
                            if (ws.Cells[4 + j, 4 + i].Value == null)
                            {
                                ws.Cells[4 + j, 4 + i].Value = -99;
                            }
                        }

                        lstItems.Add(new Item()
                        {
                            Date = Db,
                            Au = double.Parse(ws.Cells[4, 4 + i].Value.ToString()),
                            Ag = double.Parse(ws.Cells[5, 4 + i].Value.ToString()),
                            Mo = double.Parse(ws.Cells[6, 4 + i].Value.ToString()) / 10000,
                            As = double.Parse(ws.Cells[7, 4 + i].Value.ToString()),
                            Cd = double.Parse(ws.Cells[8, 4 + i].Value.ToString()),
                            Pb = double.Parse(ws.Cells[9, 4 + i].Value.ToString()),
                            Zn = double.Parse(ws.Cells[10, 4 + i].Value.ToString()),
                            Bi = double.Parse(ws.Cells[11, 4 + i].Value.ToString()),
                            Sb = double.Parse(ws.Cells[12, 4 + i].Value.ToString()),
                            FeConc = double.Parse(ws.Cells[13, 4 + i].Value.ToString()),
                            Fe = double.Parse(ws.Cells[14, 4 + i].Value.ToString()),
                            PyConc = double.Parse(ws.Cells[15, 4 + i].Value.ToString()),
                            Py = double.Parse(ws.Cells[16, 4 + i].Value.ToString()),
                            S2 = double.Parse(ws.Cells[17, 4 + i].Value.ToString()),
                            ConcentrateGrade = double.Parse(ws.Cells[18, 4 + i].Value.ToString())
                        });
                    }

                    pck.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }
                // TODO: esta esta hardcoded
                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\FreightData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        ExcelWorksheet ws2 = pck2.Workbook.Worksheets["BudgetFreight"]; // TODO: Constantes

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        int lastRow = ws2.Dimension.End.Row + 1;

                        for (int i = 0; i < lstItems.Count; i++)
                        {
                            ws2.Cells[i + lastRow, 1].Value = lstItems[i].Date;
                            ws2.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 2].Value = lstItems[i].Au;
                            ws2.Cells[i + lastRow, 3].Value = lstItems[i].Ag;
                            ws2.Cells[i + lastRow, 4].Value = lstItems[i].Mo;
                            ws2.Cells[i + lastRow, 5].Value = lstItems[i].As;
                            ws2.Cells[i + lastRow, 6].Value = lstItems[i].Cd;
                            ws2.Cells[i + lastRow, 7].Value = lstItems[i].Pb;
                            ws2.Cells[i + lastRow, 8].Value = lstItems[i].Zn;
                            ws2.Cells[i + lastRow, 9].Value = lstItems[i].Bi;
                            ws2.Cells[i + lastRow, 10].Value = lstItems[i].Sb;
                            ws2.Cells[i + lastRow, 11].Value = lstItems[i].FeConc;
                            ws2.Cells[i + lastRow, 12].Value = lstItems[i].Fe;
                            ws2.Cells[i + lastRow, 13].Value = lstItems[i].PyConc;
                            ws2.Cells[i + lastRow, 14].Value = lstItems[i].Py;
                            ws2.Cells[i + lastRow, 15].Value = lstItems[i].S2;
                            ws2.Cells[i + lastRow, 16].Value = lstItems[i].ConcentrateGrade;
                        }

                        byte[] fileText2 = pck2.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText2);

                        UpdateB = $"{StringResources.Updated}: {DateTime.Now}";

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Upload Error"); // TODO: Constantes
                    }
                }

            }
        }
    }
}
