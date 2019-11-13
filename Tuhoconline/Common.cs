using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anh.Tuhoconline
{
    public class Common
    {
        public static string GetTemplate(string asTemplateName)
        {
            string path = Path.Combine(Application.StartupPath, "Excel", asTemplateName);
            return path;
        }

        public static void SaveExcelTemplate(SpreadsheetGear.IWorkbook workbook, string fileName, string fileType,out string outPath)
        {
            string path;
            if (fileType == "xlsx")
            {
                path = Path.Combine(Path.GetTempPath(), fileName + DateTime.Now.ToString("yyMMdd_hhmmss") + ".xlsx");
                workbook.SaveAs(path, SpreadsheetGear.FileFormat.OpenXMLWorkbook);
            }
            else
            {
                path = Path.Combine(Path.GetTempPath(), fileName + DateTime.Now.ToString("yyMMdd_hhmmss") + ".xls");
                workbook.SaveAs(path, SpreadsheetGear.FileFormat.Excel8);
            }
            outPath = path;
        }
    }
}
