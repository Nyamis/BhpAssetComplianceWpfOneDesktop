using BhpAssetComplianceWpfOneDesktop.Resources;
using BhpAssetComplianceWpfOneDesktop.Utility;
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
using BhpAssetComplianceWpfOneDesktop.Constants;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Models.GeotechnicalNotesModels;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class GeotechnicalViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.Geotechnical;
        protected override string MyPosterIcon { get; set; } = IconKeys.Geotechnics;

        private string _myEscondidaImage;
        public string MyEscondidaImage
        {
            get { return _myEscondidaImage; }
            set { SetProperty(ref _myEscondidaImage, value); }
        }

        private string _myEscondidaNorteImage;
        public string MyEscondidaNorteImage
        {
            get { return _myEscondidaNorteImage; }
            set { SetProperty(ref _myEscondidaNorteImage, value); }
        }

        private ImageSource _myEscondidaImageSource;
        public ImageSource MyEscondidaImageSource
        {
            get { return _myEscondidaImageSource; }
            set { SetProperty(ref _myEscondidaImageSource, value); }
        }

        private ImageSource _myEscondidaNorteImageSource;
        public ImageSource MyEscondidaNorteImageSource
        {
            get { return _myEscondidaNorteImageSource; }
            set { SetProperty(ref _myEscondidaNorteImageSource, value); }
        }

        private bool _isEnabledLoadEscondidaImagePath;
        public bool IsEnabledLoadEscondidaImagePath
        {
            get { return _isEnabledLoadEscondidaImagePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaImagePath, value); }
        }

        private bool _isEnabledLoadEscondidaNorteImagePath;
        public bool IsEnabledLoadEscondidaNorteImagePath
        {
            get { return _isEnabledLoadEscondidaNorteImagePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaNorteImagePath, value); }
        }

        private string _myEscondidaTable;
        public string MyEscondidaTable
        {
            get { return _myEscondidaTable; }
            set { SetProperty(ref _myEscondidaTable, value); }
        }

        private string _myEscondidaNorteTable;
        public string MyEscondidaNorteTable
        {
            get { return _myEscondidaNorteTable; }
            set { SetProperty(ref _myEscondidaNorteTable, value); }
        }

        private ImageSource _myEscondidaTableSource;
        public ImageSource MyEscondidaTableSource
        {
            get { return _myEscondidaTableSource; }
            set { SetProperty(ref _myEscondidaTableSource, value); }
        }

        private ImageSource _myEscondidaNorteTableSource;
        public ImageSource MyEscondidaNorteTableSource
        {
            get { return _myEscondidaNorteTableSource; }
            set { SetProperty(ref _myEscondidaNorteTableSource, value); }
        }

        private bool _isEnabledLoadEscondidaTablePath;
        public bool IsEnabledLoadEscondidaTablePath
        {
            get { return _isEnabledLoadEscondidaTablePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaTablePath, value); }
        }

        private bool _isEnabledLoadEscondidaNorteTablePath;
        public bool IsEnabledLoadEscondidaNorteTablePath
        {
            get { return _isEnabledLoadEscondidaNorteTablePath; }
            set { SetProperty(ref _isEnabledLoadEscondidaNorteTablePath, value); }
        }

        private bool _isEnabledGenerateTemplate;
        public bool IsEnabledGenerateTemplate
        {
            get { return _isEnabledGenerateTemplate; }
            set { SetProperty(ref _isEnabledGenerateTemplate, value); }
        }

        private string _myLastRefreshValues;
        public string MyLastRefreshValues
        {
            get { return _myLastRefreshValues; }
            set { SetProperty(ref _myLastRefreshValues, value); }
        }

        private string _myLastDateRefreshImages;
        public string MyLastDateRefreshImages
        {
            get { return _myLastDateRefreshImages; }
            set { SetProperty(ref _myLastDateRefreshImages, value); }
        }

        private DateTime _myDateActual;
        public DateTime MyDateActual
        {
            get { return _myDateActual; }
            set { SetProperty(ref _myDateActual, value); }
        }
        
        public DelegateCommand SelectEscondidaImageCommand { get; private set; }
        public DelegateCommand SelectEscondidaNorteImageCommand { get; private set; }
        public DelegateCommand SelectEscondidaTableCommand { get; private set; }
        public DelegateCommand SelectEscondidaNorteTableCommand { get; private set; }
        public DelegateCommand LoadImagesCommand { get; private set; }
        public DelegateCommand GenerateGeotechnicalNotesTemplateCommand { get; private set; }
        public DelegateCommand LoadGeotechnicalNotesTemplateCommand { get; private set; }

        private readonly List<GeotechnicalNotesNotes> _notes = new List<GeotechnicalNotesNotes>();

        public GeotechnicalViewModel()
        {
            IsEnabledLoadEscondidaImagePath = false;
            IsEnabledLoadEscondidaNorteImagePath = false;
            IsEnabledLoadEscondidaTablePath = false;
            IsEnabledLoadEscondidaNorteTablePath = false;
            IsEnabledGenerateTemplate = false;
            MyDateActual = DateTime.Now;
            SelectEscondidaImageCommand = new DelegateCommand(EscondidaImagePath);
            SelectEscondidaNorteImageCommand = new DelegateCommand(EscondidaNorteImagePath);
            SelectEscondidaTableCommand = new DelegateCommand(EscondidaTablePath);
            SelectEscondidaNorteTableCommand = new DelegateCommand(EscondidaNorteTablePath);
            LoadImagesCommand = new DelegateCommand(LoadImages,CanProcess).ObservesProperty(() => IsEnabledLoadEscondidaImagePath).ObservesProperty(() => IsEnabledLoadEscondidaNorteImagePath).ObservesProperty(() => IsEnabledLoadEscondidaTablePath).ObservesProperty(() => IsEnabledLoadEscondidaNorteTablePath);
            GenerateGeotechnicalNotesTemplateCommand = new DelegateCommand(GenerateGeotechnicalNotesTemplate);
            LoadGeotechnicalNotesTemplateCommand = new DelegateCommand(LoadGeotechnicalNotesTemplate).ObservesCanExecute(() => IsEnabledGenerateTemplate);
        }

        private void EscondidaImagePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaImage = openFileDialog.FileName;
                MyEscondidaImageSource = new BitmapImage(new Uri(MyEscondidaImage));
            }
            IsEnabledLoadEscondidaImagePath = true;
        }

        private void EscondidaNorteImagePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaNorteImage = openFileDialog.FileName;
                MyEscondidaNorteImageSource = new BitmapImage(new Uri(MyEscondidaNorteImage));
            }
            IsEnabledLoadEscondidaNorteImagePath = true;
        }

        private void EscondidaTablePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaTable = openFileDialog.FileName;
                MyEscondidaTableSource = new BitmapImage(new Uri(MyEscondidaTable));
            }
            IsEnabledLoadEscondidaTablePath = true;
        }

        private void EscondidaNorteTablePath()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "All supported graphics|*.jpg;*.jpeg;*.png|" +
              "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" +
              "Portable Network Graphic (*.png)|*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MyEscondidaNorteTable = openFileDialog.FileName;
                MyEscondidaNorteTableSource = new BitmapImage(new Uri(MyEscondidaNorteTable));
            }
            IsEnabledLoadEscondidaNorteTablePath = true;
        }

        private bool CanProcess()
        {
            if (IsEnabledLoadEscondidaImagePath & IsEnabledLoadEscondidaNorteImagePath & IsEnabledLoadEscondidaTablePath & IsEnabledLoadEscondidaNorteTablePath)
            {
                return true;
            }
            return false;
        }

        private void LoadImages()
        {           
            var targetFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesCSVFilePath;
            var loadFileInfo = new FileInfo(targetFilePath);
            if (loadFileInfo.Exists)
            {
                if (targetFilePath.Substring(targetFilePath.Length - 28) == "GeotechnicalPicturesData.csv")
                {
                    try
                    {
                        var openWriteCheck = File.OpenWrite(targetFilePath);
                        openWriteCheck.Close();
                        var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);

                        var findImageOnDate = ExportImageToCsv.SearchByDateGeotechnical("Pictures", newDate, targetFilePath);
                        if (findImageOnDate == -1)
                        {
                            ExportImageToCsv.AppendImageGeotechnicalToCSV(MyEscondidaImage, MyEscondidaNorteImage, "Pictures", newDate, targetFilePath);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(targetFilePath, findImageOnDate);
                            ExportImageToCsv.AppendImageGeotechnicalToCSV(MyEscondidaImage, MyEscondidaNorteImage, "Pictures", newDate, targetFilePath);
                        }

                        var findTableOnDate = ExportImageToCsv.SearchByDateGeotechnical("Tables", newDate, targetFilePath);
                        if (findTableOnDate == -1)
                        {
                            ExportImageToCsv.AppendImageGeotechnicalToCSV(MyEscondidaTable, MyEscondidaNorteTable, "Tables", newDate, targetFilePath);
                        }
                        else
                        {
                            ExportImageToCsv.RemoveItem(targetFilePath, findTableOnDate);
                            ExportImageToCsv.AppendImageGeotechnicalToCSV(MyEscondidaTable, MyEscondidaNorteTable, "Tables", newDate, targetFilePath);
                        }
                        MyLastDateRefreshImages = $"{StringResources.Updated}: {DateTime.Now}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, StringResources.UploadError);
                    }
                }
                else
                {
                    var wrongFileMessage = $"{StringResources.WorksheetNotExist} {targetFilePath} {StringResources.IsTheRightOne}";
                    MessageBox.Show(wrongFileMessage, StringResources.UploadError);
                }               
            }
            else
            {
                var wrongFileMessage = $"{StringResources.WorksheetNotExist} {targetFilePath} {StringResources.ExistsOrNotSelect}";
                MessageBox.Show(wrongFileMessage, StringResources.UploadError);
            }           
        }

        private void GenerateGeotechnicalNotesTemplate()
        {
            var headers = new List<string>() { "Note Type", "Phase", "State", "Note" };
            var excelPackage = new ExcelPackage();

            excelPackage.Workbook.Properties.Author = "BHP";
            excelPackage.Workbook.Properties.Title = GeotechnicalNotesConstants.GeotechnicalNotesWorksheetTitle;
            excelPackage.Workbook.Properties.Company = "BHP";
            var worksheet = excelPackage.Workbook.Worksheets.Add(GeotechnicalNotesConstants.EscondidaGeotechnicalNotesWorksheet);

            for (var i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(GeotechnicalNotesTemplateColors.HeaderBackgroundGeotechnicalNotes));
                worksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (var i = 0; i < 100; i++)
            {
                for (var j = 0; j < headers.Count; j++)
                {
                    worksheet.Cells[i + 1, j + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[i + 1, j + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[i + 1, j + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[i + 1, j + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                var noteTypeListDataValidation = worksheet.Cells[i + 2, 1].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteTypeListDataValidation.AllowBlank = false;
                noteTypeListDataValidation.Formula.Values.Add("Ira");
                noteTypeListDataValidation.Formula.Values.Add("FcDc");
                noteTypeListDataValidation.ShowErrorMessage = true;
                noteTypeListDataValidation.Error = "Select from List of Values ...";

                var noteStateListDataValidations = worksheet.Cells[i + 2, 3].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteStateListDataValidations.AllowBlank = false;
                noteStateListDataValidations.Formula.Values.Add("Positive");
                noteStateListDataValidations.Formula.Values.Add("Negative");
                noteStateListDataValidations.Formula.Values.Add("Neutral");
                noteStateListDataValidations.ShowErrorMessage = true;
                noteStateListDataValidations.Error = "Select from List of Values ...";
            }

            var worksheet2 = excelPackage.Workbook.Worksheets.Add(GeotechnicalNotesConstants.EscondidaNorteGeotechnicalNotesWorksheet);

            for (var i = 0; i < headers.Count; i++)
            {
                worksheet2.Cells[1, i + 1].Value = headers[i];
                worksheet2.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet2.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet2.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(GeotechnicalNotesTemplateColors.HeaderBackgroundGeotechnicalNotes));
                worksheet2.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            for (var i = 0; i < 100; i++)
            {
                for (var j = 0; j < headers.Count; j++)
                {
                    worksheet2.Cells[i + 1, j + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 1, j + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 1, j + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 1, j + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                var noteTypeListDataValidation = worksheet2.Cells[i + 2, 1].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteTypeListDataValidation.AllowBlank = false;
                noteTypeListDataValidation.Formula.Values.Add("Ira");
                noteTypeListDataValidation.Formula.Values.Add("FcDc");
                noteTypeListDataValidation.ShowErrorMessage = true;
                noteTypeListDataValidation.Error = "Select from List of Values ...";

                var noteStateListDataValidations = worksheet2.Cells[i + 2, 3].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                noteStateListDataValidations.AllowBlank = false;
                noteStateListDataValidations.Formula.Values.Add("Positive");
                noteStateListDataValidations.Formula.Values.Add("Negative");
                noteStateListDataValidations.Formula.Values.Add("Neutral");
                noteStateListDataValidations.ShowErrorMessage = true;
                noteStateListDataValidations.Error = "Select from List of Values ...";
            }

            worksheet.Column(4).Width = 100;
            worksheet2.Column(4).Width = 100;

            byte[] fileText = excelPackage.GetAsByteArray();

            var dialog = new SaveFileDialog()
            {
                FileName = GeotechnicalNotesConstants.GeotechnicalNotesExcelFileName,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            try
            {
                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllBytes(dialog.FileName, fileText);
                    IsEnabledGenerateTemplate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, StringResources.UploadError);
            }
        }

        private void LoadGeotechnicalNotesTemplate()
        {
            _notes.Clear();
            var openFileDialog = new OpenFileDialog
            {
                Title = StringResources.SelectFile,
                Filter = "Excel Worksheets (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var openFilePath = new FileInfo(openFileDialog.FileName);
                var excelPackage = new ExcelPackage(openFilePath);
                var escondidaWorksheet = excelPackage.Workbook.Worksheets[GeotechnicalNotesConstants.EscondidaGeotechnicalNotesWorksheet];
                var escondidaNorteWorksheet = excelPackage.Workbook.Worksheets[GeotechnicalNotesConstants.EscondidaNorteGeotechnicalNotesWorksheet];

                if (openFilePath.FullName.Substring(openFilePath.FullName.Length - GeotechnicalNotesConstants.GeotechnicalNotesExcelFileName.Length) == GeotechnicalNotesConstants.GeotechnicalNotesExcelFileName)
                {
                    try
                    {
                        // Check if the file is already open
                        var openWriteCheck = File.OpenWrite(openFileDialog.FileName);
                        openWriteCheck.Close();

                        var rows = escondidaWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows; i++)
                        {
                            if (escondidaWorksheet.Cells[i + 1, 1].Value != null)
                            {
                                if (escondidaWorksheet.Cells[i + 1, 2].Value == null)
                                    escondidaWorksheet.Cells[i + 1, 2].Value = " ";
                                if (escondidaWorksheet.Cells[i + 1, 3].Value == null)
                                    escondidaWorksheet.Cells[i + 1, 3].Value = "Neutral";
                                if (escondidaWorksheet.Cells[i + 1, 4].Value == null)
                                    escondidaWorksheet.Cells[i + 1, 4].Value = " ";

                                _notes.Add(new GeotechnicalNotesNotes()
                                {
                                    Place = "Escondida Pit",
                                    NoteType = escondidaWorksheet.Cells[i + 1, 1].Value.ToString(),
                                    Phase = escondidaWorksheet.Cells[i + 1, 2].Value.ToString(),
                                    State = escondidaWorksheet.Cells[i + 1, 3].Value.ToString(),
                                    Note = escondidaWorksheet.Cells[i + 1, 4].Value.ToString()
                                });
                            }
                        }

                        var rows2 = escondidaNorteWorksheet.Dimension.Rows;
                        for (var i = 1; i < rows2; i++)
                        {
                            if (escondidaNorteWorksheet.Cells[i + 1, 1].Value != null)
                            {
                                if (escondidaNorteWorksheet.Cells[i + 1, 2].Value == null)
                                    escondidaNorteWorksheet.Cells[i + 1, 2].Value = " ";
                                if (escondidaNorteWorksheet.Cells[i + 1, 3].Value == null)
                                    escondidaNorteWorksheet.Cells[i + 1, 3].Value = "Neutral";
                                if (escondidaNorteWorksheet.Cells[i + 1, 4].Value == null)
                                    escondidaNorteWorksheet.Cells[i + 1, 4].Value = " ";

                                _notes.Add(new GeotechnicalNotesNotes()
                                {
                                    Place = "Escondida Norte Pit",
                                    NoteType = escondidaNorteWorksheet.Cells[i + 1, 1].Value.ToString(),
                                    Phase = escondidaNorteWorksheet.Cells[i + 1, 2].Value.ToString(),
                                    State = escondidaNorteWorksheet.Cells[i + 1, 3].Value.ToString(),
                                    Note = escondidaNorteWorksheet.Cells[i + 1, 4].Value.ToString()
                                });
                            }
                        }
                        excelPackage.Dispose();

                        var loadFilePath = BhpAssetComplianceWpfOneDesktop.Resources.FilePaths.Default.GeotechnicalNotesExcelFilePath;
                        var loadFileInfo = new FileInfo(loadFilePath);

                        if (loadFileInfo.Exists)
                        {
                            var package = new ExcelPackage(loadFileInfo);
                            var worksheet = package.Workbook.Worksheets[GeotechnicalNotesConstants.NotesGeotechnicalNotesSpotfireWorksheet];

                            if (worksheet != null)
                            {
                                var newDate = new DateTime(MyDateActual.Year, MyDateActual.Month, 1, 00, 00, 00);
                                var lastRow = worksheet.Dimension.End.Row + 1;

                                for (var i = 0; i < _notes.Count; i++)
                                {
                                    worksheet.Cells[i + lastRow, 1].Value = newDate;
                                    worksheet.Cells[i + lastRow, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                                    worksheet.Cells[i + lastRow, 2].Value = _notes[i].Place;
                                    worksheet.Cells[i + lastRow, 3].Value = _notes[i].NoteType;
                                    worksheet.Cells[i + lastRow, 4].Value = _notes[i].Phase;
                                    worksheet.Cells[i + lastRow, 5].Value = _notes[i].State;

                                    if (_notes[i].State == "Positive")
                                    {
                                        worksheet.Cells[i + lastRow, 6].Value = 0;
                                    }
                                    else if (_notes[i].State == "Negative")
                                    {
                                        worksheet.Cells[i + lastRow, 6].Value = 1;
                                    }
                                    else if (_notes[i].State == "Neutral")
                                    {
                                        worksheet.Cells[i + lastRow, 6].Value = 2;
                                    }
                                    worksheet.Cells[i + lastRow, 7].Value = _notes[i].Note;
                                }
                                byte[] fileText = package.GetAsByteArray();
                                File.WriteAllBytes(loadFilePath, fileText);
                                MyLastRefreshValues = $"{StringResources.Updated}: {DateTime.Now}";
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

                
            }
        }       
    }
}
