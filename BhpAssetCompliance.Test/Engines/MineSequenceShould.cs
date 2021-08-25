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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [TestMethod]
        public void QuartersReconciliationFactorsLoadTemplateFixture()
        {
            var path = @"..\..\..\TestData\FilledTemplates\ReconciliationFactorsTemplate.xlsx";
            var exists =File.Exists(path);

            var reconciliationEngine = new QuartersReconciliationFactorsReadTemplate(path);
            
        }
    }
}