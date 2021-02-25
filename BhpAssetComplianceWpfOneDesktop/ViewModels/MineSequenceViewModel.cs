using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
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
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows.Media;
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class MineSequenceViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.MineSequence;
        protected override string MyPosterIcon { get; set; } = IconKeys.MineSequence;

        public string generateContent { get; set; } = StringResources.GenerateTemplate;

        public string loadContent { get; set; } = StringResources.LoadTemplate;

        public string loadImageContent { get; set; } = StringResources.LoadImage;

        public string dateContent { get; set; } = StringResources.Date;

        public string processContent { get; set; } = StringResources.Process;

        private string _imageE;
        public string imageE
        {
            get { return _imageE; }
            set { SetProperty(ref _imageE, value); }
        }

        private string _imageEN;
        public string imageEN
        {
            get { return _imageEN; }
            set { SetProperty(ref _imageEN, value); }
        }

        ImageSource _Source;
        public ImageSource Source
        {
            get { return _Source; }
            set { SetProperty(ref _Source, value); }
        }

        private bool _isEnabled;
        public bool IsEnabled
        {
            get { return _isEnabled; }
            set { SetProperty(ref _isEnabled, value); }
        }

        ImageSource _Source2;
        public ImageSource Source2
        {
            get { return _Source2; }
            set { SetProperty(ref _Source2, value); }
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

        private string _UpdateText;
        public string UpdateText
        {
            get { return _UpdateText; }
            set { SetProperty(ref _UpdateText, value); }
        }

        DateTime _Date;
        public DateTime Date
        {
            get { return _Date; }
            set { SetProperty(ref _Date, value); }
        }

        public DelegateCommand GenerarT { get; private set; }
        public DelegateCommand CargarI1 { get; private set; }
        public DelegateCommand CargarI2 { get; private set; }
        public DelegateCommand Procesar { get; private set; }
        public DelegateCommand CargarT { get; private set; }

        public MineSequenceViewModel()
        {
            Date = DateTime.Now;
            GenerarT = new DelegateCommand(GenerateTemplate);
            CargarI1 = new DelegateCommand(ImageEPath);
            CargarI2 = new DelegateCommand(ImageENPath);
            CargarT = new DelegateCommand(ReadTemplate).ObservesCanExecute(() => IsEnabled1);
            Procesar = new DelegateCommand(Process, CanProcess).ObservesProperty(() => IsEnabled).ObservesProperty(() => IsEnabled2).ObservesProperty(() => IsEnabled3);
        }

        private void GenerateTemplate()
        {
            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Mine Sequence Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("L1Expit");

            ws.Cells["A1"].Value = "Expit Budget (t)";
            ws.Cells["B1"].Value = "Expit Actual (%)";
            ws.Cells["C1"].Value = "Budget Baseline";
            ws.Cells["A1:C1"].Style.Font.Bold = true;
            ws.Cells["A1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFC269"));
            ws.Cells["C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FAA762"));

            for (int i = 1; i < 3; i++)
            {
                ws.Cells[$"A{i}:C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[$"A{i}:C{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Column(i).Width = 18;
            }

            ws.Column(3).Width = 18;

            var ws4 = pck.Workbook.Worksheets.Add("Comments");
            ws4.Cells["A1"].Value = "Tag";
            ws4.Cells["B1"].Value = "Comment";
            ws4.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws4.Cells["A1:B1"].Style.Font.Bold = true;
            ws4.Column(1).Width = 15;
            ws4.Column(2).Width = 150;
            ws4.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws4.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#5C9FCC"));

            for (int i = 1; i < 21; i++)
            {
                ws4.Cells[$"A{i}:B{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws4.Cells[$"A{i}:B{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //var l = ws4.Cells[i + 1, 1].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                //l.AllowBlank = false;
                //l.Formula.Values.Add("Ira");
                //l.Formula.Values.Add("FcDc");
                //l.ShowErrorMessage = true;
                //l.Error = "Select from List of Values ...";
            }

            var ws2 = pck.Workbook.Worksheets.Add("AdherenceToB01L1");
            List<string> lstHeader2 = new List<string>() { "Unplanned Delay (t)", "Volume Ytd (%)", "Spatial Ytd (%)", "AdherenceL1 Ytd (%)" };

            for (int i = 0; i < lstHeader2.Count; i++)
            {
                ws2.Cells[1, i + 1].Value = lstHeader2[i];
                ws2.Cells[1, i + 1].Style.Font.Bold = true;
                ws2.Column(1 + i).Width = 21;
                ws2.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
            }
            ws2.Cells["B1:D1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            ws2.Cells["A1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D0CECE"));
            ws2.Cells["B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B62212"));
            ws2.Cells["C1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#3C4F70"));
            ws2.Cells["D1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#1A7842"));

            for (int i = 1; i < 3; i++)
            {
                ws2.Cells[$"A{i}:D{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[$"A{i}:D{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            var ws3 = pck.Workbook.Worksheets.Add("DelayRecover");
            List<string> lstHeader3 = new List<string>() { "Ytd PushBack (t)", "Phase Name", "DelayRecover Pushback (t)" };

            for (int i = 0; i < lstHeader3.Count; i++)
            {
                ws3.Cells[1, i + 1].Value = lstHeader3[i];
                ws3.Cells[1, i + 1].Style.Font.Bold = true;
                ws3.Column(1 + i).Width = 27;
                ws3.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[1, i + 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
                ws3.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#0C637C"));
            }

            ws3.Cells[$"A1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws3.Cells[$"A2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            for (int i = 1; i < 11; i++)
            {
                ws3.Cells[$"B{i}:C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws3.Cells[$"A{i}:C{i}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "MineSequenceTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                File.WriteAllBytes(dialog.FileName, fileText);
            }
            IsEnabled1 = true;
        }

        private void ImageEPath()
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
                imageE = op.FileName;
                Source = new BitmapImage(new Uri(imageE));
            }
            IsEnabled = true;
        }
        private void ImageENPath()
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
                imageEN = op.FileName;
                Source2 = new BitmapImage(new Uri(imageEN));
            }
            IsEnabled2 = true;
        }

        public class L1Expit
        {
            public double? ExpitBudgetTonnes { get; set; }
            public double? ExpitActualPercent { get; set; }
            public string BudgetBaseline { get; set; }
        }

        readonly List<L1Expit> lstL1Expit = new List<L1Expit>();

        public class AdherenceToB01L1
        {
            public double UnplannedDelayTonnes { get; set; }
            public double VolumeYtdPercent { get; set; }
            public double SpatialYtdPercent { get; set; }
            public double AdherenceL1YtdPercent { get; set; }
        }

        readonly List<AdherenceToB01L1> lstAdherenceToB01L1 = new List<AdherenceToB01L1>();

        public class DelayRecover
        {
            public double YtdPushBackTonnes { get; set; }
            public string PhaseName { get; set; }
            public double DelayRecoverPushbackTonnes { get; set; }
        }

        readonly List<DelayRecover> lstDelayRecover = new List<DelayRecover>();

        public class Comments
        {
            public string Tag { get; set; }
            public string Comment { get; set; }
        }

        readonly List<Comments> lstComments = new List<Comments>();

        private void ReadTemplate()
        {
            lstL1Expit.Clear();
            lstAdherenceToB01L1.Clear();
            lstDelayRecover.Clear();
            lstComments.Clear();

            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (op.ShowDialog() == true)
            {
                FileInfo FilePath = new FileInfo(op.FileName);
                ExcelPackage pck = new ExcelPackage(FilePath);
                ExcelWorksheet ws1 = pck.Workbook.Worksheets["L1Expit"];
                ExcelWorksheet ws2 = pck.Workbook.Worksheets["AdherenceToB01L1"];
                ExcelWorksheet ws3 = pck.Workbook.Worksheets["DelayRecover"];
                ExcelWorksheet ws4 = pck.Workbook.Worksheets["Comments"];

                try
                {
                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();

                    if (ws1.Cells[2, 1].Value == null)
                    {
                        ws1.Cells[2, 1].Value = -99;
                    }
                    if (ws1.Cells[2, 2].Value == null)
                    {
                        ws1.Cells[2, 2].Value = -99;
                    }
                    if (ws1.Cells[2, 3].Value == null)
                    {
                        ws1.Cells[2, 3].Value = " ";
                    }
                    lstL1Expit.Add(new L1Expit()
                    {
                        ExpitBudgetTonnes = double.Parse(ws1.Cells[2, 1].Value.ToString()),
                        ExpitActualPercent = double.Parse(ws1.Cells[2, 2].Value.ToString()),
                        BudgetBaseline = ws1.Cells[2, 3].Value.ToString()
                    });


                    if (ws2.Cells[2, 1].Value == null)
                    {
                        ws2.Cells[2, 1].Value = -99;
                    }
                    if (ws2.Cells[2, 2].Value == null)
                    {
                        ws2.Cells[2, 2].Value = -99;
                    }
                    if (ws2.Cells[2, 3].Value == null)
                    {
                        ws2.Cells[2, 3].Value = -99;
                    }
                    if (ws2.Cells[2, 4].Value == null)
                    {
                        ws2.Cells[2, 4].Value = -99;
                    }
                    lstAdherenceToB01L1.Add(new AdherenceToB01L1()
                    {
                        UnplannedDelayTonnes = double.Parse(ws2.Cells[2, 1].Value.ToString()),
                        VolumeYtdPercent = double.Parse(ws2.Cells[2, 2].Value.ToString()),
                        SpatialYtdPercent = double.Parse(ws2.Cells[2, 3].Value.ToString()) ,
                        AdherenceL1YtdPercent = double.Parse(ws2.Cells[2, 4].Value.ToString()) 

                    });

                    int rows = ws3.Dimension.Rows;
                    for (int i = 1; i < rows; i++)
                    {
                        if (ws3.Cells[i + 1, 2].Value != null)
                        {
                            if (ws3.Cells[2, 1].Value == null)
                            {
                                ws3.Cells[2, 1].Value = -99;
                            }
                            if (ws3.Cells[i + 1, 3].Value == null)
                            {
                                ws3.Cells[i + 1, 3].Value = -99;
                            }
                            lstDelayRecover.Add(new DelayRecover()
                            {
                                YtdPushBackTonnes = double.Parse(ws3.Cells[2, 1].Value.ToString()),
                                PhaseName = ws3.Cells[i + 1, 2].Value.ToString(),
                                DelayRecoverPushbackTonnes = double.Parse(ws3.Cells[i + 1, 3].Value.ToString()),
                            });
                        }
                    }

                    int rows2 = ws4.Dimension.Rows;
                    for (int i = 1; i < rows2; i++)
                    {
                        if (ws4.Cells[i + 1, 2].Value != null)
                        {
                            string tag;

                            if (ws4.Cells[i + 1, 1].Value.ToNullSafeString() == "")
                            {
                                tag = "All";
                            }
                            else
                            {
                                tag = ws4.Cells[i + 1, 1].Value.ToString();
                            }
                            
                            lstComments.Add(new Comments()
                            {
                                Tag = tag,
                                Comment = ws4.Cells[i + 1, 2].Value.ToString()
                            });
                        }
                    }
                    IsEnabled3 = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }
            }
        }

        private void Process()
        {
            string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\MineSequenceData.xlsx";
            FileInfo filePath = new FileInfo(fileName);

            if (filePath.Exists)
            {
                try
                {
                    ExcelPackage pck = new ExcelPackage(filePath);
                    ExcelWorksheet ws1 = pck.Workbook.Worksheets["L1Expit"];
                    ExcelWorksheet ws2 = pck.Workbook.Worksheets["AdherenceToB01L1"];
                    ExcelWorksheet ws3 = pck.Workbook.Worksheets["DelayRecover"];
                    ExcelWorksheet ws4 = pck.Workbook.Worksheets["Comments"];

                    FileStream fs = File.OpenWrite(fileName);
                    fs.Close();

                    DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00);
                    int lastRow1 = ws1.Dimension.End.Row + 1;

                    ws1.Cells[lastRow1, 1].Value = newDate;
                    ws1.Cells[lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                    ws1.Cells[lastRow1, 2].Value = lstL1Expit[0].ExpitBudgetTonnes;
                    ws1.Cells[lastRow1, 3].Value = lstL1Expit[0].ExpitActualPercent;
                    ws1.Cells[lastRow1, 4].Value = lstL1Expit[0].BudgetBaseline;

                    int lastRow2 = ws2.Dimension.End.Row + 1;

                    ws2.Cells[lastRow2, 1].Value = newDate;
                    ws2.Cells[lastRow2, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                    ws2.Cells[lastRow2, 2].Value = lstAdherenceToB01L1[0].UnplannedDelayTonnes;
                    ws2.Cells[lastRow2, 3].Value = lstAdherenceToB01L1[0].VolumeYtdPercent;
                    ws2.Cells[lastRow2, 4].Value = lstAdherenceToB01L1[0].SpatialYtdPercent;
                    ws2.Cells[lastRow2, 5].Value = lstAdherenceToB01L1[0].AdherenceL1YtdPercent;

                    int lastRow3 = ws3.Dimension.End.Row + 1;

                    for (int i = 0; i < lstDelayRecover.Count; i++)
                    {
                        ws3.Cells[i + lastRow3, 1].Value = newDate;
                        ws3.Cells[i + lastRow3, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws3.Cells[i + lastRow3, 2].Value = lstDelayRecover[i].YtdPushBackTonnes;
                        ws3.Cells[i + lastRow3, 3].Value = lstDelayRecover[i].PhaseName;
                        ws3.Cells[i + lastRow3, 4].Value = lstDelayRecover[i].DelayRecoverPushbackTonnes;
                    }

                    int lastRow4 = ws4.Dimension.End.Row + 1;

                    for (int i = 0; i < lstComments.Count; i++)
                    {
                        ws4.Cells[i + lastRow4, 1].Value = newDate;
                        ws4.Cells[i + lastRow4, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws4.Cells[i + lastRow4, 2].Value = lstComments[i].Tag;
                        ws4.Cells[i + lastRow4, 3].Value = lstComments[i].Comment;
                    }

                    string target = @"c:\users\nyamis\oneDrive - bmining\BHP\MineSequencePictureData.csv";

                    int findE = ExportImageToCsv.SearchByDateMineSequence(newDate, target);
                    if (findE == -1)
                    {
                        ExportImageToCsv.AppendImageMineSequenceToCSV(imageE, imageEN, newDate, target);
                    }
                    else
                    {
                        ExportImageToCsv.RemoveItem(target, findE);
                        ExportImageToCsv.AppendImageMineSequenceToCSV(imageE, imageEN, newDate, target);
                    }

                    byte[] fileText = pck.GetAsByteArray();
                    File.WriteAllBytes(fileName, fileText);

                    UpdateText = $"{StringResources.Updated}: {DateTime.Now}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

            }

        }

        private bool CanProcess()
        {
            return IsEnabled & IsEnabled2 & IsEnabled3;
        }
    }
}
