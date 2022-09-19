using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MSExcel = Microsoft.Office.Interop.Excel;


namespace DHD.ExcelAddInTools
{
    internal class Common
    {

        public static MSExcel.Worksheet ActiveSheet { get { return (MSExcel.Worksheet)App.ActiveSheet; } }
        public static MSExcel.Application App { get { return (MSExcel.Application)Globals.ThisAddIn.Application; } }
        public static MSExcel.Workbook ActiveBook { get { return (MSExcel.Workbook)App.ActiveWorkbook; } }
        public static MSExcel.Range ActiveCell { get { return (MSExcel.Range)App.ActiveCell; } }



    }
}
