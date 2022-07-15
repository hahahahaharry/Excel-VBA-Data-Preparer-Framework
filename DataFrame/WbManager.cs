using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace DataPreparer
{
    public class WbManager
    {
        private static WbManager _instance = null;
        private const String ProcessName = "EXCEL";
        private const String MarshalName = "Excel.Application";
        
        public Int32 SessionID { get; private set; }
        

        public static WbManager GetInstance() 
        {
            if(_instance == null)
            {
                _instance = new WbManager();
            }
            return _instance;
        }

        public Workbook GetWorkbook(string workbookPath)
        {
            ExcelApp app = GetExcelInstance();
            if (File.Exists(workbookPath))
            {
                var workbookName = Path.GetFileName(workbookPath);
                if (app.Workbooks.Count > 0)
                {
                    for (int i = 1; i <= app.Workbooks.Count; i++)
                    {
                        var workbook = app.Workbooks[i];
                        if (workbook.Name == workbookName)
                        {
                            return workbook;
                        }
                    }
                }
                return app.Workbooks.Open(workbookPath);
            }
            else
            {
                return null;
            }
        }

        private ExcelApp GetExcelInstance()
        {
            try
            {
                return Marshal.GetActiveObject(MarshalName) as ExcelApp;
            }
            catch (COMException)
            {
                return new ExcelApp();
            }
        }

    }
}
