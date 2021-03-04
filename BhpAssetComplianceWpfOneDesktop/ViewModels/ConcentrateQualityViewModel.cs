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
        private string _myLastDateRefreshActualValues; // TODO: Respetar en la variables privadas la notación _updateA
        public string MyLastDateRefreshActualValues
        {
            get { return _myLastDateRefreshActualValues; }
            set { SetProperty(ref _myLastDateRefreshActualValues, value); }
        }

        private string _myLastRefreshBudgetValues;
        public string MyLastRefreshBudgetValues
        {
            get { return _myLastRefreshBudgetValues; }
            set { SetProperty(ref _myLastRefreshBudgetValues, value); }
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


        private readonly List<ConcentrateQualityFreight> _freights = new List<ConcentrateQualityFreight>();

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
            // TODO: Cambiar a headers y usar var
            var headers = new List<string> { "Nombre M/N", "N°", "Inicio embarque", "Termino embarque" };
            var items = new List<string> { "WMT", "DMT", "Moisture.", "Cu", "As", "Fe", "Au", "Ag", "S", "Insol.", "Cd", "Zn", "Hg", "SiO2", "Al2O3", "Sb", "Mo" };
            var units = new List<string> { "Pesometer t", "Pesometer t", "%", "%", "%", "%", "g/t", "g/t", "%", "%", "%", "%", "g/t", "%", "%", "%", "%" };

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = "Real Month Freight Template";
            excelPackage.Workbook.Properties.Company = "BHP";


            var worksheet = excelPackage.Workbook.Worksheets.Add(ConcentrateQualityConstants.RealFreightWorksheet);
            // TODO: usar mejores variables de números
            worksheet.Cells["A2:U2"].Style.Font.Bold = true; // TODO: Evitar usar referencias "A1" y usar números
            worksheet.Cells["A2:U2"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));  // TODO: Utilizar constantes para colores
            worksheet.Cells["E3:U3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
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
                    worksheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#375623"));
                }
                else if (i % 2 == 0)
                {
                    worksheet.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#833C0C"));
                }
                worksheet.Column(1 + i).Width = 16;
            }

            for (int i = 1; i < 18; i++)
            {
                if (i % 2 != 0)
                {
                    worksheet.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#548235"));
                }
                else if (i % 2 == 0)
                {
                    worksheet.Cells[3, 4 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[3, 4 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C65911"));
                }
            }

            worksheet.Cells["A2:U2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            for (int i = 0; i < 14; i++)
            {
                worksheet.Cells[i + 2, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{2 + i}:U{2 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = excelPackage.GetAsByteArray();

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
                    IsEnabledLoadActualValues = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Upload Error");
            }
        }


        private void LoadActualFreightTemplate()
        {
            _freights.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);

                try
                {
                    var worksheet = excelPackage.Workbook.Worksheets[ConcentrateQualityConstants.RealFreightWorksheet];
                    var fileStream = File.OpenWrite(openFileDialog.FileName);
                    fileStream.Close();

                    var rows = worksheet.Dimension.Rows;

                    for (var i = 1; i < rows; i++)
                    {
                        if (worksheet.Cells[i + 3, 1].Value != null)
                        {

                            for (int j = 0; j < 17; j++)
                                if (worksheet.Cells[3 + i, 5 + j].Value == null)
                                    worksheet.Cells[3 + i, 5 + j].Value = -99;
                            _freights.Add(new ConcentrateQualityFreight()
                            {

                                Name = worksheet.Cells[3 + i, 1].Value.ToString(),
                                Number = Int32.Parse(worksheet.Cells[3 + i, 2].Value.ToString()),
                                Start = Convert.ToDateTime(worksheet.Cells[3 + i, 3].Value.ToString()),
                                End = Convert.ToDateTime(worksheet.Cells[3 + i, 4].Value.ToString()),
                                WMT = double.Parse(worksheet.Cells[3 + i, 5].Value.ToString()),
                                DMT = double.Parse(worksheet.Cells[3 + i, 6].Value.ToString()),
                                Moisture = double.Parse(worksheet.Cells[3 + i, 7].Value.ToString()) / 100,
                                Cu = double.Parse(worksheet.Cells[3 + i, 8].Value.ToString()) / 100,
                                As = double.Parse(worksheet.Cells[3 + i, 9].Value.ToString()) * 10000,
                                Fe = double.Parse(worksheet.Cells[3 + i, 10].Value.ToString()) / 100,
                                Au = double.Parse(worksheet.Cells[3 + i, 11].Value.ToString()),
                                Ag = double.Parse(worksheet.Cells[3 + i, 12].Value.ToString()),
                                S = double.Parse(worksheet.Cells[3 + i, 13].Value.ToString()) / 100,
                                Insoluble = double.Parse(worksheet.Cells[3 + i, 14].Value.ToString()) / 100,
                                Cd = double.Parse(worksheet.Cells[3 + i, 15].Value.ToString()) * 10000,
                                Zn = double.Parse(worksheet.Cells[3 + i, 16].Value.ToString()) * 10000,
                                Hg = double.Parse(worksheet.Cells[3 + i, 17].Value.ToString()),
                                SiO2 = double.Parse(worksheet.Cells[3 + i, 18].Value.ToString()) / 100,
                                Al2O3 = double.Parse(worksheet.Cells[3 + i, 19].Value.ToString()) / 100,
                                Sb = double.Parse(worksheet.Cells[3 + i, 20].Value.ToString()) * 10000,
                                Mo = double.Parse(worksheet.Cells[3 + i, 21].Value.ToString())
                            });
                        }
                    }
                    excelPackage.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }
                // TODO: Implementar sistema de archivos default.
                var loadFilePath = @"c:\users\nyamis\oneDrive - bmining\BHP\FreightData.xlsx";
                var loadFileInfo = new FileInfo(loadFilePath);

                if (loadFileInfo.Exists)
                {
                    try
                    {
                        ExcelPackage package = new ExcelPackage(loadFileInfo);
                        ExcelWorksheet worksheet = package.Workbook.Worksheets["RealMonth"];

                        // Check if the file is already open
                        var openWriteCheck = File.OpenWrite(loadFilePath);
                        openWriteCheck.Close();

                        int lastRow = worksheet.Dimension.End.Row + 1;
                        DateTime newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00).AddMilliseconds(000);

                        for (int i = 0; i < _freights.Count; i++)
                        {
                            worksheet.Cells[i + lastRow, 1].Value = newDate;
                            worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            worksheet.Cells[i + lastRow, 2].Value = _freights[i].Name;
                            worksheet.Cells[i + lastRow, 3].Value = _freights[i].Number;
                            worksheet.Cells[i + lastRow, 4].Value = _freights[i].Start;
                            worksheet.Cells[i + lastRow, 4].Style.Numberformat.Format = "yyyy-MM-dd";
                            worksheet.Cells[i + lastRow, 5].Value = _freights[i].End;
                            worksheet.Cells[i + lastRow, 5].Style.Numberformat.Format = "yyyy-MM-dd";
                            worksheet.Cells[i + lastRow, 6].Value = _freights[i].WMT;
                            worksheet.Cells[i + lastRow, 7].Value = _freights[i].DMT;
                            worksheet.Cells[i + lastRow, 8].Value = _freights[i].Moisture;
                            worksheet.Cells[i + lastRow, 9].Value = _freights[i].Cu;
                            worksheet.Cells[i + lastRow, 10].Value = _freights[i].As;
                            worksheet.Cells[i + lastRow, 11].Value = _freights[i].Fe;
                            worksheet.Cells[i + lastRow, 12].Value = _freights[i].Au;
                            worksheet.Cells[i + lastRow, 13].Value = _freights[i].Ag;
                            worksheet.Cells[i + lastRow, 14].Value = _freights[i].S;
                            worksheet.Cells[i + lastRow, 15].Value = _freights[i].Insoluble;
                            worksheet.Cells[i + lastRow, 16].Value = _freights[i].Cd;
                            worksheet.Cells[i + lastRow, 17].Value = _freights[i].Zn;
                            worksheet.Cells[i + lastRow, 18].Value = _freights[i].Hg;
                            worksheet.Cells[i + lastRow, 19].Value = _freights[i].SiO2;
                            worksheet.Cells[i + lastRow, 20].Value = _freights[i].Al2O3;
                            worksheet.Cells[i + lastRow, 21].Value = _freights[i].Sb;
                            worksheet.Cells[i + lastRow, 22].Value = _freights[i].Mo;
                        }
                        byte[] fileText2 = package.GetAsByteArray();
                        File.WriteAllBytes(loadFilePath, fileText2);
                        MyLastDateRefreshActualValues = $"{StringResources.Updated}: {DateTime.Now}";
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
            ws.Cells["B2"].Value = $"FY{MyFiscalYear}";

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
                    IsEnabledLoadBudgetValues = true;
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
                            Db = new DateTime(MyFiscalYear - 1, M, 1, 00, 00, 00).AddMilliseconds(000); // TODO: No es necesario el add milliseconds
                        }
                        else if (i == 6 || i == 7 || i == 8 || i == 9 || i == 10 || i == 11)
                        {
                            Db = new DateTime(MyFiscalYear, M, 1, 00, 00, 00).AddMilliseconds(000);
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

                        MyLastRefreshBudgetValues = $"{StringResources.Updated}: {DateTime.Now}";

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
