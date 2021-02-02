using BhpAssetComplianceWpfOneDesktop.Models;
using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows.Media;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class GeotechnicalViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.Geotechnical;

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

        private bool _isEnabled1;
        public bool IsEnabled1
        {
            get { return _isEnabled1; }
            set { SetProperty(ref _isEnabled1, value); }
        }

        ImageSource _Source2;
        public ImageSource Source2
        {
            get { return _Source2; }
            set { SetProperty(ref _Source2, value); }
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

        public GeotechnicalViewModel()
        {
            Date = DateTime.Now;
            GenerarT = new DelegateCommand(GenerateTemplate);
            CargarT = new DelegateCommand(ReadTemplate).ObservesCanExecute(() => IsEnabled1);
            CargarI1 = new DelegateCommand(ImageEPath);
            CargarI2 = new DelegateCommand(ImageENPath);
            Procesar = new DelegateCommand(Process, CanProcess).ObservesProperty(() => IsEnabled).ObservesProperty(() => IsEnabled2).ObservesProperty(() => IsEnabled3);
        }

        private void GenerateTemplate()
        {
            List<string> lstHeader = new List<string>() { "Note Type", "Phase", "State", "Note" };

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Properties.Author = "BHP";
            pck.Workbook.Properties.Title = "Geotechnical Notes Template";
            pck.Workbook.Properties.Company = "BHP";

            var ws = pck.Workbook.Worksheets.Add("Escondida");

            //Header Section
            for (int i = 0; i < lstHeader.Count; i++)
            {
                ws.Cells[1, i + 1].Value = lstHeader[i];
                ws.Cells[1, i + 1].Style.Font.Bold = true;
                ws.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < 100; i++)
            {
                for (int j = 0; j < lstHeader.Count; j++)
                {
                    ws.Cells[i + 1, j + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[i + 1, j + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[i + 1, j + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[i + 1, j + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                var l = ws.Cells[i + 2, 1].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                l.AllowBlank = false;
                l.Formula.Values.Add("Ira");
                l.Formula.Values.Add("FcDc");
                l.ShowErrorMessage = true;
                l.Error = "Select from List of Values ...";


                var s = ws.Cells[i + 2, 3].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                s.AllowBlank = false;
                s.Formula.Values.Add("Positive");
                s.Formula.Values.Add("Negative");
                s.Formula.Values.Add("Neutral");
                s.ShowErrorMessage = true;
                s.Error = "Select from List of Values ...";

            }

            var ws2 = pck.Workbook.Worksheets.Add("Escondida Norte");

            //Header Section
            for (int i = 0; i < lstHeader.Count; i++)
            {
                ws2.Cells[1, i + 1].Value = lstHeader[i];
                ws2.Cells[1, i + 1].Style.Font.Bold = true;
                ws2.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A9D08E"));
                ws2.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (int i = 0; i < 100; i++)
            {
                for (int j = 0; j < lstHeader.Count; j++)
                {
                    ws2.Cells[i + 1, j + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws2.Cells[i + 1, j + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws2.Cells[i + 1, j + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws2.Cells[i + 1, j + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                var l = ws2.Cells[i + 2, 1].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                l.AllowBlank = false;
                l.Formula.Values.Add("Ira");
                l.Formula.Values.Add("FcDc");
                l.ShowErrorMessage = true;
                l.Error = "Select from List of Values ...";

                var s = ws2.Cells[i + 2, 3].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                s.AllowBlank = false;
                s.Formula.Values.Add("Positive");
                s.Formula.Values.Add("Negative");
                s.Formula.Values.Add("Neutral");
                s.ShowErrorMessage = true;
                s.Error = "Select from List of Values ...";
            }

            ws.Column(4).Width = 100;
            ws2.Column(4).Width = 100;

            byte[] fileText = pck.GetAsByteArray();

            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "GeotechnicalNotesTemplate.xlsx",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                File.WriteAllBytes(dialog.FileName, fileText);
            }
            IsEnabled1 = true;
        }

        public class Notes
        {
            public string Place { get; set; }
            public string NoteType { get; set; }
            public string Phase { get; set; }
            public string State { get; set; }
            public string Note { get; set; }
        }

        readonly List<Notes> lstNotes = new List<Notes>();

        private void ReadTemplate()
        {
            lstNotes.Clear();
            OpenFileDialog op = new OpenFileDialog
            {
                Title = "Select File",
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (op.ShowDialog() == true)
            {
                FileInfo FilePath = new FileInfo(op.FileName);
                ExcelPackage pck = new ExcelPackage(FilePath);
                ExcelWorksheet ws = pck.Workbook.Worksheets["Escondida"];
                ExcelWorksheet ws2 = pck.Workbook.Worksheets["Escondida Norte"];

                try
                {                   
                    FileStream fs = File.OpenWrite(op.FileName);
                    fs.Close();
                    int rows = ws.Dimension.Rows;

                    for (int i = 1; i < rows; i++)
                    {
                        if (ws.Cells[i + 1, 1].Value != null)
                        {
                            if (ws.Cells[i + 1, 2].Value == null)
                            {
                                ws.Cells[i + 1, 2].Value = " ";
                            }
                            if (ws.Cells[i + 1, 3].Value == null)
                            {
                                ws.Cells[i + 1, 3].Value = "Neutral";
                            }
                            if (ws.Cells[i + 1, 4].Value == null)
                            {
                                ws.Cells[i + 1, 4].Value = " ";
                            }
                            lstNotes.Add(new Notes()
                            {
                                Place = "Escondida Pit",
                                NoteType = ws.Cells[i + 1, 1].Value.ToString(),
                                Phase = ws.Cells[i + 1, 2].Value.ToString(),
                                State = ws.Cells[i + 1, 3].Value.ToString(),
                                Note = ws.Cells[i + 1, 4].Value.ToString()
                            });
                        }

                    }

                    int rows2 = ws2.Dimension.Rows;
                    for (int i = 1; i < rows2; i++)
                    {
                        if (ws2.Cells[i + 1, 1].Value != null)
                        {
                            if (ws2.Cells[i + 1, 2].Value == null)
                            {
                                ws2.Cells[i + 1, 2].Value = " ";
                            }
                            if (ws2.Cells[i + 1, 3].Value == null)
                            {
                                ws2.Cells[i + 1, 3].Value = "Neutral";
                            }
                            if (ws2.Cells[i + 1, 4].Value == null)
                            {
                                ws2.Cells[i + 1, 4].Value = " ";
                            }
                            lstNotes.Add(new Notes()
                            {
                                Place = "Escondida Norte Pit",
                                NoteType = ws2.Cells[i + 1, 1].Value.ToString(),
                                Phase = ws2.Cells[i + 1, 2].Value.ToString(),
                                State = ws2.Cells[i + 1, 3].Value.ToString(),
                                Note = ws2.Cells[i + 1, 4].Value.ToString()
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

        private bool CanProcess()
        {
            if (IsEnabled & IsEnabled2 & IsEnabled3)
            {
                return true;
            }
            return false;
        }

        private void Process()
        {
            string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\GeotechnicalNotesData.xlsx";
            FileInfo filePath = new FileInfo(fileName);


            if (filePath.Exists)
            {
                ExcelPackage pck = new ExcelPackage(filePath);
                ExcelWorksheet ws = pck.Workbook.Worksheets["Notes"];

                try
                {
                    FileStream fs = File.OpenWrite(fileName);
                    fs.Close();

                    DateTime newDate = new DateTime(Date.Year, Date.Month, 1, 00, 00, 00);
                    int lastRow = ws.Dimension.End.Row + 1;

                    for (int i = 0; i < lstNotes.Count; i++)
                    {
                        ws.Cells[i + lastRow, 1].Value = newDate;
                        ws.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws.Cells[i + lastRow, 2].Value = lstNotes[i].Place;
                        ws.Cells[i + lastRow, 3].Value = lstNotes[i].NoteType;
                        ws.Cells[i + lastRow, 4].Value = lstNotes[i].Phase;
                        ws.Cells[i + lastRow, 5].Value = lstNotes[i].State;

                        if (lstNotes[i].State == "Positive")
                        {
                            ws.Cells[i + lastRow, 6].Value = 0;
                        }
                        else if (lstNotes[i].State == "Negative")
                        {
                            ws.Cells[i + lastRow, 6].Value = 1;
                        }
                        else if (lstNotes[i].State == "Neutral")
                        {
                            ws.Cells[i + lastRow, 6].Value = 2;
                        }

                        ws.Cells[i + lastRow, 7].Value = lstNotes[i].Note;
                    }

                    string target = @"c:\users\nyamis\oneDrive - bmining\BHP\GeotechnicalPictureData.csv";

                    int findE = ExportImageToCsv.SearchByDate("Escondida Pit", newDate, target);
                    if (findE == -1)
                    {
                        ExportImageToCsv.AppendImageToCSV(imageE, "Escondida Pit", newDate, target);
                    }
                    else
                    {
                        ExportImageToCsv.RemoveItem(target, findE);
                        ExportImageToCsv.AppendImageToCSV(imageE, "Escondida Pit", newDate, target);
                    }

                    int findEN = ExportImageToCsv.SearchByDate("Escondida Norte Pit", newDate, target);
                    if (findEN == -1)
                    {
                        ExportImageToCsv.AppendImageToCSV(imageEN, "Escondida Norte Pit", newDate, target);
                    }
                    else
                    {
                        ExportImageToCsv.RemoveItem(target, findEN);
                        ExportImageToCsv.AppendImageToCSV(imageEN, "Escondida Norte Pit", newDate, target);
                    }

                    byte[] fileText = pck.GetAsByteArray();
                    File.WriteAllBytes(fileName, fileText);

                    UpdateText = $"Actualizado: {DateTime.Now}";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Upload Error");
                }

            }

        }
    }
}
