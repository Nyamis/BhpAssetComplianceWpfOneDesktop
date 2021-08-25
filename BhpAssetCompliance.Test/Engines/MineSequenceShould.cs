using System.IO;
using BhpAssetComplianceWpfOneDesktop.Engines;
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
            // SOLID
            // Single Purpose Class 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [TestMethod]
        public void QuartersReconciliationFactorsLoadTemplateFixture()
        {
            var path = @"..\..\..\TestData\FilledTemplates\ReconciliationFactorsTemplate.xlsx";
            var exists =File.Exists(path);

            // var reconcilationEngine = new QuarterReconciliationFactorsLoadTemplateEngine();
        }
    }
}