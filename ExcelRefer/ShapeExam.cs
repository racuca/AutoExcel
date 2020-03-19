using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelShape = Microsoft.Office.Core;

namespace ExcelRefer
{
    class ShapeExam
    {

        public ShapeExam()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            //xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Shapes.AddShape(ExcelShape.MsoAutoShapeType.msoShapeRectangle, 50, 50, 100, 100);

            xlWorkSheet.Shapes.SelectAll();
            //xlWorkSheet.Shapes.Item(0).Fill.ForeColor.RGB = Color.Red.ToArgb();

            xlWorkBook.SaveAs(System.IO.Path.Combine(Environment.CurrentDirectory, "ShapeTest.xls"), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
