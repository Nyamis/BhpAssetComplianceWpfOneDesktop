using System.IO;
using System.Linq;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using BhpAssetComplianceWpfOneDesktop.Engines;
using BhpAssetComplianceWpfOneDesktop.ViewModels;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace BhpAssetCompliance.Test.Engines
{
    [TestClass]
    public class MineSequenceShould
    {
        [TestInitialize]
        public void Test()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [TestMethod]
        public void QuartersReconciliationFactorsReadTemplateFixture()
        {
            var path = @"..\..\..\TestData\FilledTemplates\ReconciliationFactorsTemplate.xlsx";
            var reconciliationEngine = new QuartersReconciliationFactorsReadTemplate(path);
            var reconciliationFactors = reconciliationEngine.Process();
            reconciliationFactors.F1.First(f1 => f1.Quarter == "Q2").MillCuT.Should()
                .BeApproximately(105.67390626502 / 100, 0.0001);
        }


        [TestMethod]
        public void QuartersReconciliationFactorsLoadTemplateFixture()
        {
            var path = @"..\..\..\TestData\FilledTemplates\ReconciliationFactorsTemplate.xlsx";
            var reconciliationEngine = new QuartersReconciliationFactorsReadTemplate(path);
            var reconciliationFactors = reconciliationEngine.Process();


            var pathTemplateResult = @"..\..\..\TestData\FilledTemplates\ReconciliationFactorsTemplateOutput.xlsx";

            // Delete files
            var outputPath = @"..\..\..\TestData\TestResults\OutputReconciliationFactors.xlsx";
            if (File.Exists(outputPath))
                File.Delete(outputPath);
            File.Copy(pathTemplateResult, outputPath);

            var quartersReconciliationFactorsLoadTemplateEngine =
                new QuartersReconciliationFactorsLoadTemplateEngine(reconciliationFactors, outputPath);
        }
    }
}