namespace BhpAssetComplianceWpfOneDesktop.Engines
{
    public class QuartersReconciliationFactorsLoadTemplateEngine
    {
        private readonly QuarterReconciliationFactors _factors;
        private readonly string _outputPath;

        public QuartersReconciliationFactorsLoadTemplateEngine(QuarterReconciliationFactors factors, string outputPath)
        {
            _factors = factors;
            _outputPath = outputPath;
        }

        public void LoadTemplate()
        {
        }
    }
}