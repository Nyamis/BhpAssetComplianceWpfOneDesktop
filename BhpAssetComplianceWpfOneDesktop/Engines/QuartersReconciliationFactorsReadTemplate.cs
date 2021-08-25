using System;
using System.IO;
using BhpAssetComplianceWpfOneDesktop.Constants;
using BhpAssetComplianceWpfOneDesktop.Models.QuartersReconciliationFactorsModels;
using BhpAssetComplianceWpfOneDesktop.Resources;
using OfficeOpenXml;

namespace BhpAssetComplianceWpfOneDesktop.Engines
{
    public class QuartersReconciliationFactorsReadTemplate
    {
        private readonly string _path;

        public QuartersReconciliationFactorsReadTemplate(string path)
        {
            _path = path;
        }

        public QuarterReconciliationFactors Process()
        {
            var factors = new QuarterReconciliationFactors
            {
            };


            // Check if the name of the file is standard
            // TODO: Agregar el checkeo de los estilos
            if (_path.Substring(_path.Length -
                                QuartersReconciliationFactorsConstants.QuartersReconciliationFactorExcelFileName
                                    .Length) == QuartersReconciliationFactorsConstants
                .QuartersReconciliationFactorExcelFileName)
            {
                var wrongFileMessage = $"{StringResources.WrongUploadedFile} {_path} {StringResources.IsTheRightOne}";
                throw new Exception(wrongFileMessage);
            }

            // Check if the file is already open
            try
            {
                var fileStream = File.OpenWrite(_path);
                fileStream.Close();
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }

            // Start gathering the information
            var excelPackage = new ExcelPackage(new FileInfo(_path));
            var worksheet =
                excelPackage.Workbook.Worksheets[
                    QuartersReconciliationFactorsConstants.QuartersReconciliationFactorsWorksheet];


            for (var i = 0; i < 4; i++)
            {
                // Check TODO:
                for (var j = 0; j < 10; j++)
                    if (worksheet.Cells[7 + i, 4 + j].Value == null)
                        worksheet.Cells[7 + i, 4 + j].Value = -99;
                // Check TODO:
                for (var j = 0; j < 9; j++)
                    if (worksheet.Cells[18 + i, 4 + j].Value == null)
                        worksheet.Cells[18 + i, 4 + j].Value = -99;
                // Check TODO:
                for (var j = 0; j < 6; j++)
                    if (worksheet.Cells[29 + i, 4 + j].Value == null)
                        worksheet.Cells[29 + i, 4 + j].Value = -99;
            }

            for (var i = 0; i < 4; i++)
            {
                factors.F0.Add(new QuartersReconciliationFactorsF0()
                {
                    Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
                    MillOre = double.Parse(worksheet.Cells[7 + i, 4].Value.ToString()) / 100,
                    OLOre = double.Parse(worksheet.Cells[18 + i, 4].Value.ToString()) / 100,
                    SLOre = double.Parse(worksheet.Cells[29 + i, 4].Value.ToString()) / 100,
                    MillCuT = double.Parse(worksheet.Cells[7 + i, 5].Value.ToString()) / 100,
                    OLCuT = double.Parse(worksheet.Cells[18 + i, 5].Value.ToString()) / 100,
                    SLCuT = double.Parse(worksheet.Cells[29 + i, 5].Value.ToString()) / 100,
                    MillCuFines = double.Parse(worksheet.Cells[7 + i, 6].Value.ToString()) / 100,
                    OLCuFines = double.Parse(worksheet.Cells[18 + i, 6].Value.ToString()) / 100,
                    SLCuFines = double.Parse(worksheet.Cells[29 + i, 6].Value.ToString()) / 100
                });
            }

            for (var i = 0; i < 4; i++)
            {
                factors.F1.Add(new QuartersReconciliationFactorsF1()
                {
                    Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
                    MillOre = double.Parse(worksheet.Cells[7 + i, 7].Value.ToString()) / 100,
                    OLOre = double.Parse(worksheet.Cells[18 + i, 7].Value.ToString()) / 100,
                    SLOre = double.Parse(worksheet.Cells[29 + i, 7].Value.ToString()) / 100,
                    MillCuT = double.Parse(worksheet.Cells[7 + i, 8].Value.ToString()) / 100,
                    OLCuT = double.Parse(worksheet.Cells[18 + i, 8].Value.ToString()) / 100,
                    SLCuT = double.Parse(worksheet.Cells[29 + i, 8].Value.ToString()) / 100,
                    MillCuFines = double.Parse(worksheet.Cells[7 + i, 9].Value.ToString()) / 100,
                    OLCuFines = double.Parse(worksheet.Cells[18 + i, 9].Value.ToString()) / 100,
                    SLCuFines = double.Parse(worksheet.Cells[29 + i, 9].Value.ToString()) / 100
                });
            }

            for (var i = 0; i < 4; i++)
            {
                factors.F2.Add(new QuartersReconciliationFactorsF2()
                {
                    Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
                    MillOre = double.Parse(worksheet.Cells[7 + i, 10].Value.ToString()) / 100,
                    OLOre = double.Parse(worksheet.Cells[18 + i, 10].Value.ToString()) / 100,
                    MillCuT = double.Parse(worksheet.Cells[7 + i, 11].Value.ToString()) / 100,
                    OLCuT = double.Parse(worksheet.Cells[18 + i, 11].Value.ToString()) / 100,
                    MillCuFines = double.Parse(worksheet.Cells[7 + i, 12].Value.ToString()) / 100,
                    OLCuFines = double.Parse(worksheet.Cells[18 + i, 12].Value.ToString()) / 100
                });
            }

            for (var i = 0; i < 4; i++)
            {
                factors.F3.Add(new QuartersReconciliationFactorsF3()
                {
                    Quarter = worksheet.Cells[7 + i, 3].Value.ToString(),
                    MillCuFines = double.Parse(worksheet.Cells[7 + i, 13].Value.ToString()) / 100
                });
            }
            excelPackage.Dispose();
            return factors;
        }
    }
}