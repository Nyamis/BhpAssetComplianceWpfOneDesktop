using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BhpAssetComplianceWpfOneDesktop.Utility
{
    public static class TemplateDates
    {
        public static DateTime ConvertDateToFiscalYearDate(int iterator, int fiscalYear, int month)
        {
            var _date = DateTime.Now;

            if (iterator == 0 || iterator == 1 || iterator == 2 || iterator == 3 || iterator == 4 || iterator == 5)
            {
                _date = new DateTime(fiscalYear - 1, month, 1, 00, 00, 00);
                return _date;
            }
            else /*if (iterator == 6 || iterator == 7 || iterator == 8 || iterator == 9 || iterator == 10 || iterator == 11)*/
            {
                _date = new DateTime(fiscalYear, month, 1, 00, 00, 00);
                return _date;
            }
        }

        public static string ConvertDateToFiscalYearString(DateTime date)
        {
            var _year = $"{date.Year}";
            var _fiscalYear = Int32.Parse(_year.Substring((_year.Length - 2), 2));
            if (date.Month == 7 || date.Month == 8 || date.Month == 9 || date.Month == 10 || date.Month == 11 || date.Month == 12)
            {
                _fiscalYear = _fiscalYear + 1;
                var fiscalYearString = $"FY{_fiscalYear}";
                return fiscalYearString;
            }
            else
            {
                var fiscalYearString = $"FY{_fiscalYear}";
                return fiscalYearString;
            }
        }

    }
}
