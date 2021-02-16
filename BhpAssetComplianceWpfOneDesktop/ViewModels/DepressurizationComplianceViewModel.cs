using BhpAssetComplianceWpfOneDesktop.Models;
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

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class DepressurizationComplianceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.DepressurizationCompliance;

        public string generateContent { get; set; } = StringResources.GenerateTemplate;
        public string loadContent { get; set; } = StringResources.LoadTemplate;
        public string dateContent { get; set; } = StringResources.Date;
        public string fiscalYearContent { get; set; } = StringResources.FiscalYear;
        public string monthlyContent { get; set; } = StringResources.MonthlyCompliancedepressurizationTemplate;
        public string targetContent { get; set; } = StringResources.TargetDepressurizationTemplate;
        public string loadImageContent { get; set; } = StringResources.LoadImage;


        private string _image;
        public string image
        {
            get { return _image; }
            set { SetProperty(ref _image, value); }
        }

        ImageSource _Source;
        public ImageSource Source
        {
            get { return _Source; }
            set { SetProperty(ref _Source, value); }
        }

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

        private bool _isEnabled3;
        public bool IsEnabled3
        {
            get { return _isEnabled3; }
            set { SetProperty(ref _isEnabled3, value); }
        }

        public DelegateCommand CargarI1 { get; private set; }
        public DelegateCommand GenerarDMT { get; private set; }
        public DelegateCommand CargarDMT { get; private set; }
        public DelegateCommand GenerarDTT { get; private set; }
        public DelegateCommand CargarDTT { get; private set; }

        public DepressurizationComplianceViewModel()
        {
            Date = DateTime.Now;
            FiscalYear = Date.Year;
            CargarI1 = new DelegateCommand(ImagePath);
            GenerarDMT = new DelegateCommand(GenerateDepressurizationMonthlyTemplate);
            CargarDMT = new DelegateCommand(LoadDepressurizationMonthlyTemplate, CanLoadDMT).ObservesProperty(() => IsEnabled1).ObservesProperty(() => IsEnabled2);
            GenerarDTT = new DelegateCommand(GenerateDepressurizationTargetTemplate);
            CargarDTT = new DelegateCommand(LoadDepressurizationTargetTemplate).ObservesCanExecute(() => IsEnabled3);            
        }

        private void ImagePath()
        {
            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (op.ShowDialog() == true)
            {
                image = op.FileName;
                Source = new BitmapImage(new Uri(op.FileName));
            }
            IsEnabled1 = true;
        }

        private void GenerateDepressurizationMonthlyTemplate()
        {
            List<string> lstHeader = new List<string>() { "Zona", "Observado (MPa)", "Compliance (%)", "Pit" };
            List<string> lstZone = new List<string>() { "Pared Noreste Fuera Rajo", "Pared Noreste", "Pared Noreste Talud Bajo", "Pared Los Colorados", "Pared Los Colorados Talud Bajo", "Pared Este Fuera Rajo", "Pared Este Talud Medio" };
            ExcelPackage pck = new ExcelPackage();

            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Monthly Compliance Depressurization Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("MonthlyDepressurization");

            for (int i = 0; i < lstHeader.Count; i++)
            {
                ws.Cells[1, i + 1].Value = lstHeader[i];
                ws.Cells[1, i + 1].Style.Font.Bold = true;
                ws.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(1 + i).Width = 16;
                ws.Cells[1, 1 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, 1 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9E1F2"));
            }
            ws.Column(1).Width = 27;

            for (int i = 0; i < lstZone.Count; i++)
            {
                ws.Cells[2 + i, 1].Value = lstZone[i];
            }

            for (int i = 0; i <= 12; i++)
            {
                ws.Cells[i + 1, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"A{1 + i}:D{1 + i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "DepressurizationMonthlyTemplate.xlsx",
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

        public class MonthlyCompliance
        {
            public string Zone { get; set; }
            public double Observado { get; set; }
            public double Compliance { get; set; }
            public string Pit { get; set; }
        }

        readonly List<MonthlyCompliance> lstCompliance = new List<MonthlyCompliance>();

        private bool CanLoadDMT()
        {
            if (IsEnabled1 & IsEnabled2)
            {
                return true;
            }
            return false;
        }

        private void LoadDepressurizationMonthlyTemplate()
        {

            lstCompliance.Clear();
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
                    ExcelWorksheet ws = pck.Workbook.Worksheets["MonthlyDepressurization"];

                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    int rows = ws.Dimension.Rows;

                    for (int i = 1; i < rows; i++)
                    {
                        if (ws.Cells[i + 1, 1].Value != null)
                        {
                            for (int j = 0; j < 2; j++)
                            {
                                if (ws.Cells[1 + i, 2 + j].Value == null)
                                {
                                    ws.Cells[1 + i, 2 + j].Value = -99;
                                }
                            }

                            if (ws.Cells[1 + i, 4].Value == null)
                            {
                                ws.Cells[1 + i, 4].Value = " ";
                            }

                            lstCompliance.Add(new MonthlyCompliance()
                            {
                                Zone = ws.Cells[1 + i, 1].Value.ToString(),
                                Observado = double.Parse(ws.Cells[1 + i, 2].Value.ToString()),
                                Compliance = double.Parse(ws.Cells[1 + i, 3].Value.ToString()),
                                Pit = ws.Cells[1 + i, 4].Value.ToString()
                            });
                        }
                    }
                    pck.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\DepressurizationComplianceData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        ExcelWorksheet ws2 = pck2.Workbook.Worksheets["Observado"];

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();
                        DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00);
                        int lastRow = ws2.Dimension.End.Row + 1;

                        for (int i = 0; i < lstCompliance.Count; i++)
                        {
                            ws2.Cells[i + lastRow, 1].Value = newDate;
                            ws2.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 2].Value = lstCompliance[i].Zone;
                            ws2.Cells[i + lastRow, 3].Value = lstCompliance[i].Observado;
                            ws2.Cells[i + lastRow, 4].Value = lstCompliance[i].Compliance;
                            ws2.Cells[i + lastRow, 5].Value = lstCompliance[i].Pit;
                        }

                        string target = @"c:\users\nyamis\oneDrive - bmining\BHP\DepressurizationCompliancePictureData.csv";

                        int findI = ExportImageToCsv.SearchByDateMineSequence(newDate, target);
                        if (findI == -1)
                        {
                            ExportImageToCsv.AppendImageDepressurizationToCSV(image, newDate, target);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(target, findI);
                            ExportImageToCsv.AppendImageDepressurizationToCSV(image, newDate, target);
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

        private void GenerateDepressurizationTargetTemplate()
        {
            List<string> lstMonths = new List<string>() { "July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June" };
            List<string> lstPlace = new List<string>() { "Pared Noreste Fuera Rajo", "Pared Noreste", "Pared Noreste Talud Bajo", "Pared Los Colorados", "Pared Los Colorados Talud Bajo", "Pared Este Fuera Rajo", "Pared Este Talud Medio" };

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Target Depressurization Template";
            pck.Workbook.Properties.Company = "BHP";
            var ws = pck.Workbook.Worksheets.Add("Target Compliance");
            ws.Protection.IsProtected = true;

            for (int i = 0; i < 14; i++)
            {
                ws.Column(1 + i).Style.Locked = false;
            }

            ws.Cells["A2:A3"].Merge = true;
            ws.Cells["A2"].Value = "Depressurization Place";
            ws.Cells["A2:B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A2:B2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B2"].Style.Font.Bold = true;
            ws.Column(1).Style.Font.Bold = true;
            ws.Column(1).Width = 27;
            for (int i = 0; i < lstPlace.Count; i++)
            {
                ws.Cells[4 + i, 1].Value = lstPlace[i];
            }
            ws.Cells["A4:A13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A4:A13"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9E1F2"));

            ws.Cells["B2:M2"].Merge = true;
            ws.Cells["B2"].Value = $"FY{FiscalYear}";
            ws.Cells["A2:A3"].Style.Font.Bold = true;
            ws.Cells["A2:M2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A2:M2"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9E1F2"));

            ws.Row(3).Style.Font.Bold = true;
            for (int i = 0; i < lstMonths.Count; i++)
            {
                ws.Cells[3, 2 + i].Value = lstMonths[i];
                ws.Cells[3, 2 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[3, 2 + i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9E1F2"));
                ws.Cells[3, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(2 + i).Width = 10;
            }

            for (int i = 0; i < 12; i++)
            {
                ws.Cells[$"A{1 + i}:M{1 + i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                for (int j = 0; j < 13; j++)
                {
                    ws.Cells[2 + i, 1 + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
            }
            ws.Cells[$"A13:M13"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "DepressurizationTargetTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                FileStream fs = File.OpenWrite(dialog.FileName);
                fs.Close();
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabled3 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Upload Error");
            }
        }

        public class TargetD
        {
            public DateTime Date { get; set; }
            public string Zone { get; set; }
            public double Target { get; set; }
        }

        readonly List<TargetD> lstTarget = new List<TargetD>();

        private void LoadDepressurizationTargetTemplate()
        {
            lstTarget.Clear();
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
                    ExcelWorksheet ws = pck.Workbook.Worksheets["Target Compliance"];

                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    DateTime Db = DateTime.Now;

                    for (int i = 0; i < 12; i++)
                    {
                        int M = DateTime.ParseExact(ws.Cells[3, 2 + i].Value.ToString(), "MMMM", CultureInfo.InvariantCulture).Month;

                        if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i == 5)
                        {
                            Db = new DateTime(FiscalYear - 1, M, 1, 00, 00, 00).AddMilliseconds(000);
                        }
                        else if (i == 6 || i == 7 || i == 8 || i == 9 || i == 10 || i == 11)
                        {
                            Db = new DateTime(FiscalYear, M, 1, 00, 00, 00).AddMilliseconds(000);
                        }

                        int rows = ws.Dimension.Rows;

                        for (int j = 0; j < rows; j++)
                        {
                            if (ws.Cells[4 + j, 1].Value != null)
                            {
                                if (ws.Cells[4 + j, 2 + i].Value == null)
                                {
                                    ws.Cells[4 + j, 2 + i].Value = -99;
                                }

                                lstTarget.Add(new TargetD()
                                {
                                    Date = Db,
                                    Zone = ws.Cells[4 + j, 1].Value.ToString(),
                                    Target = double.Parse(ws.Cells[4 + j, 2 + i].Value.ToString())
                                });
                            }
                        }

                    }

                    pck.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\DepressurizationComplianceData.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {                  
                    try
                    {
                        ExcelPackage pck2 = new ExcelPackage(filePath);
                        ExcelWorksheet ws2 = pck2.Workbook.Worksheets["Target"];

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        int lastRow = ws2.Dimension.End.Row + 1;

                        for (int i = 0; i < lstTarget.Count; i++)
                        {
                            ws2.Cells[i + lastRow, 1].Value = lstTarget[i].Date;
                            ws2.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                            ws2.Cells[i + lastRow, 2].Value = lstTarget[i].Zone;
                            ws2.Cells[i + lastRow, 3].Value = lstTarget[i].Target;
                        }

                        byte[] fileText2 = pck2.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText2);

                        UpdateB = $"{StringResources.Updated}: {DateTime.Now}";
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
