using Application.Services.ValueObjects;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Services.ExcelService
{
    public class ExcelService : IDisposable
    {
        private Excel.Application _application;
        private Workbooks _workbooks;
        private Workbook _workbook;
        private Worksheet _worksheet;

        private ExcelService(Excel.Application application, Workbooks workbooks, Workbook workbook, Worksheet worksheet)
        {
            _application = application;
            _workbooks = workbooks;
            _workbook = workbook;
            _worksheet = worksheet;
        }

        public void Dispose()
        {
            _workbook.Close(true);
            while (Marshal.ReleaseComObject(_workbook) != 0);

            _workbook = null;
            _workbooks.Close();
            while (Marshal.ReleaseComObject(_workbooks) != 0);
            while (Marshal.ReleaseComObject(_worksheet) != 0);

            _worksheet = null;
            _workbooks = null;
            _application.Quit();
            
            while(Marshal.ReleaseComObject(_application) != 0);

            _application = null;

            GC.Collect();
            Thread.Sleep(200);
        }

        public void WriteToCell(ExcelCellCoords coords, string value)
        {
            var cell = _worksheet.Cells[coords.Vertical, coords.Horizontal];
            cell.Value = value;
            Marshal.ReleaseComObject(cell);
        }
        public static ExcelService OpenExcelFile(string filePath)
        {
            Excel.Application application = null;
            Workbooks workbooks = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                application = new();
                workbooks = application.Workbooks;
                workbook = workbooks.Open(filePath
                    , 0, false, 5, "",
                    "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    "\t", true, false,
                    0, true, 1, 0);
                worksheet = workbook.ActiveSheet;

                return new ExcelService(application, workbooks, workbook, worksheet);
            }
            catch (Exception)
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                }
                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }
                throw;
            }
        }
    }
}
