using System.Drawing;
using BhpAssetComplianceWpfOneDesktop.Constants.TemplateColors;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BhpAssetComplianceWpfOneDesktop.Extensions
{
    public static class ExcelExtensions
    {
        public static void SetAllBorders(this ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
        }

        public static void SetMainHeader(this ExcelRange range, Color color,
            ExcelBorderStyle borderStyle = ExcelBorderStyle.Thick, bool isBold = true)
        {
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
            range.Style.Font.Bold = isBold;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }
    }
}