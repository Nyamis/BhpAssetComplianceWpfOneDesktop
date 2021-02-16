using System;
using System.Collections.Generic;
using BhpAssetComplianceWpfOneDesktop.Models;
using BhpAssetComplianceWpfOneDesktop.Resources;
using Prism.Commands;
using System.Windows;
using OfficeOpenXml;
using System.IO;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public class HistoricalRecordViewModel : BasePosterViewModel
    {
        protected override string MyPosterName { get; set; } = StringResources.HistoricalRecord;

        public string dateContent { get; set; } = StringResources.Date;
        public string areaContent { get; set; } = StringResources.Area;
        public string commentContent { get; set; } = StringResources.Comment;
        public string loadContent { get; set; } = StringResources.LoadRecord;

        private string _update;
        public string update
        {
            get { return _update; }
            set { SetProperty(ref _update, value); }
        }

        DateTime _date;
        public DateTime date
        {
            get { return _date; }
            set { SetProperty(ref _date, value); }
        }

        private string _comment;
        public string comment
        {
            get { return _comment; }
            set { SetProperty(ref _comment, value); }
        }

        private List<String> _lstArea;
        public List<String> lstArea
        {
            get
            {
                return new List<string>() { StringResources.MineSequence, StringResources.MineCompliance, StringResources.DepressurizationCompliance, StringResources.Geotechnical, StringResources.QuartersReconciliationFactors, StringResources.ProcessCompliance, StringResources.Other };
            }
            set
            {
                _lstArea = value;
            }
        }

        private string _area;
        public string area
        {
            get { return _area; }
            set { SetProperty(ref _area, value); }
        }


        public DelegateCommand LoadRecord { get; private set; }

        public HistoricalRecordViewModel()
        {
            date = DateTime.Now;
            LoadRecord = new DelegateCommand(LoadNewRecord);
        }

        private void LoadNewRecord()
        {
            if (comment != null)
            {
                string fileName = @"c:\users\nyamis\oneDrive - bmining\BHP\AssetComplianceHistoricalRecord.xlsx";
                FileInfo filePath = new FileInfo(fileName);

                if (filePath.Exists)
                {
                    try
                    {
                        ExcelPackage pck = new ExcelPackage(filePath);
                        ExcelWorksheet ws = pck.Workbook.Worksheets["Record"];

                        FileStream fs = File.OpenWrite(fileName);
                        fs.Close();

                        DateTime newDate = new DateTime(date.Year, date.Month, 1, 00, 00, 00);
                        int lastRow1 = ws.Dimension.End.Row + 1;

                        ws.Cells[lastRow1, 1].Value = newDate;
                        ws.Cells[lastRow1, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                        ws.Cells[lastRow1, 2].Value = area;
                        ws.Cells[lastRow1, 3].Value = comment;

                        byte[] fileText = pck.GetAsByteArray();
                        File.WriteAllBytes(fileName, fileText);

                        update = $"{StringResources.Updated}: {DateTime.Now}";
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
