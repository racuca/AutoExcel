﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using AutoExcel.Shapes;

namespace AutoExcel.Core
{
    public class ExcelBinder : IDisposable
    {
        public const string UID = "Excel.Application";
        object oExcel = null;
        object WorkBooks, WorkBook, WorkSheets, WorkSheet; //, Interior, Cells;
        int excelversion = 2007;

                
        public enum XlPasteType
        {
            // 요약:
            //     Only the values will be pasted.
            xlPasteValues = -4163,
            //
            // 요약:
            //     Comments will be pasted.
            xlPasteComments = -4144,
            //
            // 요약:
            //     Formulas will be pasted.
            xlPasteFormulas = -4123,
            //
            // 요약:
            //     Formatting will be pasted.
            xlPasteFormats = -4122,
            //
            // 요약:
            //     Everything will be pasted.
            xlPasteAll = -4104,
            //
            // 요약:
            //     Validation from the source cell is applied to the destination cell.
            xlPasteValidation = 6,
            //
            // 요약:
            //     Everything except borders will be pasted.
            xlPasteAllExceptBorders = 7,
            //
            // 요약:
            //     The column width of the source cell will be applied to the destination cell.
            xlPasteColumnWidths = 8,
            //
            // 요약:
            //     Formulas and number formats are pasted.
            xlPasteFormulasAndNumberFormats = 11,
            //
            // 요약:
            //     Only the values number formats will be pasted.
            xlPasteValuesAndNumberFormats = 12,
            //
            // 요약:
            //     Everything will be pasted using the source theme.
            xlPasteAllUsingSourceTheme = 13,
            xlPasteAllMergingConditionalFormats = 14
        }

        public ExcelBinder()
        {            
            Type classType = Type.GetTypeFromProgID(UID);
            oExcel = Activator.CreateInstance(classType);

            MessageFilter.Register();

            string ver = version;

            if (ver.StartsWith("11."))
            {
                excelversion = 2003;
            }
            else if (ver.StartsWith("12."))
            {
                excelversion = 2007;
            }
            else if (ver.StartsWith("14."))
            {
                excelversion = 2010;
            }
            else if (ver.StartsWith("15."))
            {
                excelversion = 2013;
            }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(WorkSheet);
            Marshal.ReleaseComObject(WorkSheets);
            Marshal.ReleaseComObject(WorkBook);
            Marshal.ReleaseComObject(WorkBooks);
            Marshal.ReleaseComObject(oExcel);
            //GC.GetTotalMemory(true);
        }


        public string version
        {
            get
            {
                return oExcel.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, oExcel, null).ToString();
            }
        }

        public bool Visible
        {
            set
            {
                oExcel.GetType().InvokeMember("Visible", BindingFlags.SetProperty,
                    null, oExcel, new object[] { value });
            }
            get
            {
                return Convert.ToBoolean(oExcel.GetType().InvokeMember("Visible", BindingFlags.GetProperty,
                   null, oExcel, null));
            }
        }

        public bool DisplayAlerts
        {
            set
            {
                oExcel.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, oExcel, new Object[] { value });
            }
            get
            {
                return Convert.ToBoolean(oExcel.GetType().InvokeMember("DisplayAlerts", BindingFlags.GetProperty, null, oExcel, null));
            }
        }

        public bool DisplayScrollBarsVisible
        {
            set
            {
                oExcel.GetType().InvokeMember("DisplayScrollBars", BindingFlags.SetProperty,
                    null, oExcel, new object[] { value });
            }
            get
            {
                return Convert.ToBoolean(oExcel.GetType().InvokeMember("DisplayScrollBars", BindingFlags.GetProperty,
                   null, oExcel, null));
            }
        }

        public bool DisplayStatusBarVisible
        {
            set
            {
                oExcel.GetType().InvokeMember("DisplayStatusBar", BindingFlags.SetProperty,
                    null, oExcel, new object[] { value });
            }
            get
            {
                return Convert.ToBoolean(oExcel.GetType().InvokeMember("DisplayStatusBar", BindingFlags.GetProperty,
                   null, oExcel, null));
            }
        }

        public string Caption
        {
            set
            {
                oExcel.GetType().InvokeMember("Caption", BindingFlags.SetProperty,
                    null, oExcel, new object[] { value });
            }
            get
            {
                return Convert.ToString(oExcel.GetType().InvokeMember("Caption", BindingFlags.GetProperty,
                    null, oExcel, null));
            }
        }

        public string Version
        {
            get
            {
                return oExcel.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, oExcel, null).ToString();
            }
        }

        public bool ScreenUpdating
        {
            get
            {
                return Convert.ToBoolean(oExcel.GetType().InvokeMember(
                                                "ScreenUpdating", BindingFlags.GetProperty, 
                                                null, oExcel, new object[] { false })
                                        );
            }
            set
            {
                oExcel.GetType().InvokeMember("ScreenUpdating", BindingFlags.SetProperty, 
                                                null, oExcel, new object[] { false }
                                             );
            }
        }




        public enum XlWindowState
        {
            xlMaximized = -4137,
            xlMinimized = -4140,
            xlNormal = -4143
        }

        public XlWindowState WindowState
        {
            set
            {
                oExcel.GetType().InvokeMember("WindowState", BindingFlags.SetProperty,
                    null, oExcel, new object[] { value });
            }
        }

        // Open Existing excel file
        public void OpenDocument(string name)
        {
            WorkBooks = oExcel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, oExcel, null);
            WorkBook = WorkBooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, WorkBooks, new object[] { name });
            WorkSheets = WorkBook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, WorkBook, null);
            WorkSheet = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { 1 });
        }

        // Create new excel file
        public void NewDocument()
        {
            WorkBooks = oExcel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, oExcel, null);
            WorkBook = WorkBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, WorkBooks, null);
            WorkSheets = WorkBook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, WorkBook, null);
            WorkSheet = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { 1 });
        }

        // Save editing excel file
        public void SaveDocument(string name)
        {
            if (File.Exists(name))
                WorkBook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null,
                    WorkBook, null);
            else
                WorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null,
                    WorkBook, new object[] { name });
        }

        // Save as editing excel file
        public void SaveAsDocument(string name)
        {
            WorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null,
                                        WorkBook, new object[] { name });
        }


        public void CloseDocument()
        {
            object result = WorkBook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, WorkBook, new object[] { true });
        }

        public void QuitDocument()
        {
            object result = oExcel.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oExcel, null);

            // and turn off the IOleMessageFilter.
            MessageFilter.Revoke();
        }

        public void SetWorkSheet(int i)
        {
            WorkSheet = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { i });
        }

        public void SetActiveWorkSheet(string name)
        {
            int count = GetSheetCount();
            for (int i = 1; i <= count; i++)
            {
                object sheetobj = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { i });
                string sheetName = sheetobj.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sheetobj, null).ToString();
                if (name.Equals(sheetName))
                {
                    sheetobj.GetType().InvokeMember("Activate", BindingFlags.GetProperty, null, sheetobj, null);
                    Marshal.ReleaseComObject(sheetobj);
                    WorkSheet = WorkBook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, WorkBook, null);
                    break;
                }
                Marshal.ReleaseComObject(sheetobj);
            }
        }

        public void SetActiveWorkSheet(int idx)
        {
            object sheetobj = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { idx });
            sheetobj.GetType().InvokeMember("Activate", BindingFlags.GetProperty, null, sheetobj, null);
            Marshal.ReleaseComObject(sheetobj);
            WorkSheet = WorkBook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, WorkBook, null);
        }

        public object GetActiveWorkSheet()
        {
            WorkSheet = WorkBook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, WorkBook, null);
            return WorkSheet;
        }

        public int GetSheetCount()
        {
            return (int)WorkSheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, WorkSheets, null);
        }

        public string GetSheetName(int SheetNumber)
        {
            WorkSheet = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { SheetNumber });
            return (string)WorkSheet.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, WorkSheet, null);
        }

        public int GetSheetRow()
        {
            object Cells = WorkSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, WorkSheet, null);
            object Range = Cells.GetType().InvokeMember("SpecialCells", BindingFlags.GetProperty, null, Cells, new object[] { 11 });

            return (int)Range.GetType().InvokeMember("Row", BindingFlags.GetProperty, null, Range, null); ;
        }

        public void SetColor(string range, int color)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });

            object Interior = Range.GetType().InvokeMember("Interior", BindingFlags.GetProperty,
                null, Range, null);

            Range.GetType().InvokeMember("Color", BindingFlags.SetProperty, null,
                Interior, new object[] { color });
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Print Functions
        // ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public enum XlPageOrientation
        {
            xlPortrait = 1,
            xlLandscape = 2
        }

        public void SetOrientation(XlPageOrientation Orientation)
        {
            object PageSetup = WorkSheet.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty,
                null, WorkSheet, null);

            PageSetup.GetType().InvokeMember("Orientation", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Orientation });
        }

        public void SetMargin(double Left, double Right, double Top, double Bottom)
        {
            object PageSetup = WorkSheet.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty,
                null, WorkSheet, null);

            PageSetup.GetType().InvokeMember("LeftMargin", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Left });
            PageSetup.GetType().InvokeMember("RightMargin", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Right });
            PageSetup.GetType().InvokeMember("TopMargin", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Top });
            PageSetup.GetType().InvokeMember("BottomMargin", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Bottom });
        }

        public enum xlPaperSize
        {
            xlPaperA4 = 9,
            xlPaperA4Small = 10,
            xlPaperA5 = 11,
            xlPaperLetter = 1,
            xlPaperLetterSmall = 2,
            xlPaper10x14 = 16,
            xlPaper11x17 = 17,
            xlPaperA3 = 9,
            xlPaperB4 = 12,
            xlPaperB5 = 13,
            xlPaperExecutive = 7,
            xlPaperFolio = 14,
            xlPaperLedger = 4,
            xlPaperLegal = 5,
            xlPaperNote = 18,
            xlPaperQuarto = 15,
            xlPaperStatement = 6,
            xlPaperTabloid = 3
        }

        public void SetPaperSize(xlPaperSize Size)
        {
            object PageSetup = WorkSheet.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty,
                null, WorkSheet, null);

            PageSetup.GetType().InvokeMember("PaperSize", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Size });
        }

        public void SetZoom(int Percent)
        {
            object PageSetup = WorkSheet.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty,
                null, WorkSheet, null);

            PageSetup.GetType().InvokeMember("Zoom", BindingFlags.SetProperty,
                null, PageSetup, new object[] { Percent });
        }

        enum PrintMode
        {
            AllPrint = 0,
            OneSheet
        }


        PrintOption option;
        public void SetPrintOption(PrintOption option)
        {
            this.option = option;
        }

        //public void Print(PrintMode m, int sheetNum, int num)
        //{
        //    bool prevDisplayAlerts = DisplayAlerts;
        //    bool prevScreenUpdating = ScreenUpdating;

        //    int sheetcount = GetSheetCount();

        //    for (int i = 1; i <= sheetcount; i++)
        //    {
        //        SetActiveWorkSheet(i);
        //        SetOrientation(option.pageShape);
        //        SetPaperSize(option.papersize);
        //        WorkSheet.GetType().InvokeMember("PrintOut", BindingFlags.SetProperty, null, WorkSheet, null);
        //    }
        //}
/*
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = "Data" Then
                ActiveSheet.Next.Select
            End If
            If ws.Name <> "Data" Then
            '   Print 1st Range - Portrait
                With ActiveSheet.PageSetup
                    .PrintArea = "$B$31:$J$90"
                    .Orientation = xlPortrait
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                    ActiveWindow.SelectedSheets.PrintOut copies:=1, Collate:=True
            '   Print 2nd Range - Landscape
                With ActiveSheet.PageSetup
                    .PrintArea = "$S$91:$AJ$128"
                    .Orientation = xlLandscape
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                    ActiveWindow.SelectedSheets.PrintOut copies:=1, Collate:=True
            End If
                Range("A1").Select
            Next ws
            
        Sheets("Data").Select

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub
*/

        public void RenameSheet(int n, string Name)
        {
            object Page = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { n });

            Page.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, Page, new object[] { Name });
        }

        public void AddNewSheet(string Name)
        {
            SetActiveWorkSheet(1);

            WorkSheet = WorkSheets.GetType().InvokeMember("Add", BindingFlags.GetProperty, null, WorkSheets, null);
            object Page = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { 1 });
            Page.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, Page, new object[] { Name });

            // Move to Last index
            int count = GetSheetCount();
            object sheetAfter = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { count });
            Page.GetType().InvokeMember("Move", BindingFlags.InvokeMethod, null, Page, new object[] { Missing.Value, sheetAfter });

            Marshal.ReleaseComObject(Page);
            Marshal.ReleaseComObject(sheetAfter);
        }

        public void AddNewSheet(string Name, int pos)
        {
            WorkSheet = WorkSheets.GetType().InvokeMember("Add", BindingFlags.GetProperty, null, WorkSheets, null);

            object sheetbefore = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { pos });
            WorkSheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, WorkSheet, new object[] { Name });
            WorkSheet.GetType().InvokeMember("Move", BindingFlags.InvokeMethod, null, WorkSheet, new object[] { sheetbefore, Missing.Value });
        }
        
        public void SetFont(string range, Font font)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });

            object Font = Range.GetType().InvokeMember("Font", BindingFlags.GetProperty,
                null, Range, null);

            Range.GetType().InvokeMember("Name", BindingFlags.SetProperty, null,
                Font, new object[] { font.Name });

            Range.GetType().InvokeMember("Size", BindingFlags.SetProperty, null,
                Font, new object[] { font.Size });

            Marshal.FinalReleaseComObject(Range);
        }
        
        public string GetRangeStr(int row, int col)
        {
            return GetExcelColumnName(col) + row.ToString();
        }

        public void SetValue(string range, string value)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, Range, new object[] { value });

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetValue(int row, int col, object value)
        {
            string range = GetRangeStr(row, col);
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, Range, new object[] { value });

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetValues(string range, object[,] value)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("Value2", BindingFlags.SetProperty, null, Range, new object[] { value });

            Marshal.FinalReleaseComObject(Range);
        }


        public void SetMerge(string range, bool MergeCells)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, Range, new object[] { MergeCells });

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetColumnWidth(string range, double Width)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Width };
            Range.GetType().InvokeMember("ColumnWidth", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void EntireColumnAutoFit(string startCell, string endCell)
        {
            object[] args = new object[] { startCell, endCell };

            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet, args);
            object EntireColumn = Range.GetType().InvokeMember("EntireColumn", BindingFlags.GetProperty, null, Range, null);

            EntireColumn.GetType().InvokeMember("AutoFit", BindingFlags.InvokeMethod, null, EntireColumn, null);

            Marshal.FinalReleaseComObject(Range);
        }

        public void EntireColumnAutoFit(int startrow, int endrow, int startcol, int endcol)
        {
            object startCell = GetCell(startrow, endrow);
            object endCell = GetCell(endrow, endcol);

            object[] args = new object[] { startCell, endCell };

            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                    null, WorkSheet, args);

            object EntireColumn = Range.GetType().InvokeMember("EntireColumn", BindingFlags.GetProperty, null, Range, null);

            EntireColumn.GetType().InvokeMember("AutoFit", BindingFlags.SetProperty, null, EntireColumn, null);

            Marshal.FinalReleaseComObject(Range);
        }


        public void SetTextOrientation(string range, int Orientation)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Orientation };
            Range.GetType().InvokeMember("Orientation", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetVerticalAlignment(string range, int Alignment)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Alignment };
            Range.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetHorisontalAlignment(string range, int Alignment)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Alignment };
            Range.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SelectText(string range, int Start, int Length, int Color, string FontStyle, int FontSize)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Start, Length };
            object Characters = Range.GetType().InvokeMember("Characters", BindingFlags.GetProperty, null, Range, args);
            object Font = Characters.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, Characters, null);
            Font.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, Font, new object[] { Color });
            Font.GetType().InvokeMember("FontStyle", BindingFlags.SetProperty, null, Font, new object[] { FontStyle });
            Font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, Font, new object[] { FontSize });

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetWrapText(string range, bool Value)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Value };
            Range.GetType().InvokeMember("WrapText", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void CreateComment(string range, bool CommentVisible, string Text)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("AddComment", BindingFlags.InvokeMethod, null, Range, null);
            object Comment = Range.GetType().InvokeMember("Comment", BindingFlags.GetProperty, null, Range, null);
            Comment.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Comment, new object[] { false });
            Comment.GetType().InvokeMember("Text", BindingFlags.InvokeMethod, null, Comment, new object[] { Text });

            Marshal.FinalReleaseComObject(Range);
        }

        public void DeleteComment(string range)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("ClearComments", BindingFlags.InvokeMethod, null, Range, null);

            Marshal.FinalReleaseComObject(Range);
        }

        public enum XlCommentDisplayMode
        {
            xlCommentAndIndicator = 1,
            xlCommentIndicatorOnly = -1,
            xlNoIndicator = 0
        }

        public void DisplayCommentIndicator(XlCommentDisplayMode Mode)
        {
            //Application.DisplayCommentIndicator
            oExcel.GetType().InvokeMember("DisplayCommentIndicator", BindingFlags.SetProperty,
                null, oExcel, new object[] { Mode });
        }

        public void SetRowsGroup(string range, bool Value)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object Rows = Range.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, Range, null);
            if (Value)
                Rows.GetType().InvokeMember("Group", BindingFlags.GetProperty, null, Rows, null);
            else
                Rows.GetType().InvokeMember("Ungroup", BindingFlags.GetProperty, null, Rows, null);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetColumnsGroup(string range, bool Value)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object Columns = Range.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, Range, null);
            if (Value)
                Columns.GetType().InvokeMember("Group", BindingFlags.GetProperty, null, Columns, null);
            else
                Columns.GetType().InvokeMember("Ungroup", BindingFlags.GetProperty, null, Columns, null);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetRowHeight(string range, double Height)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { Height };
            Range.GetType().InvokeMember("RowHeight", BindingFlags.SetProperty, null, Range, args);

            Marshal.FinalReleaseComObject(Range);
        }

        public void SetBorderStyle(string range, int Style)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            object[] args = new object[] { 1 };
            object[] args1 = new object[] { 1 };
            object Borders = Range.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, Range, null);
            Borders = Range.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, Borders, args);

            Marshal.FinalReleaseComObject(Range);
            Marshal.FinalReleaseComObject(Borders);
        }

        public object GetCell(int row, int col)
        {
            object cells = WorkSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, WorkSheet, null);
            object range = cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, new object[] { row, col });
            return range;
        }

        public object GetRow(int row)
        {
            object result = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet,
                new object[] { GetCell(row, 1), GetCell(row, 256) });
            return result;
        }

        public void SelectCell(int row, int col)
        {
            object cell = GetCell(row, col);
            cell.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, cell, null);
        }

        public string GetValue(int row, int col)
        {
            string range = GetExcelColumnName(col) + row.ToString();
            return GetValue(range);
        }

        public string GetValue(string range)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet, new object[] { range });

            object val;
            val = Range.GetType().InvokeMember("Value", BindingFlags.GetProperty,
                null, Range, null);

            Marshal.FinalReleaseComObject(Range);

            if (val == null)
                return null;

            return val.ToString();
        }

        public double GetPosition(string range)
        {
            object Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            double pos = double.Parse(Range.GetType().InvokeMember("Left", BindingFlags.GetProperty,
                null, Range, null).ToString());

            Marshal.FinalReleaseComObject(Range);

            return pos;
        }

        public bool SheetExists(string name)
        {
            int count = GetSheetCount();
            for (int i = 1; i <= count; i++)
            {
                object sheetobj = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { i });
                string sheetName = sheetobj.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sheetobj, null).ToString();
                if (name.Equals(sheetName))
                    return true;
            }
            return false;
        }

        public void DeleteSheet(string name)
        {
            int count = GetSheetCount();
            for (int i = 1; i <= count; i++)
            {
                object sheetobj = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { i });
                string sheetName = sheetobj.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sheetobj, null).ToString();
                if (name.Equals(sheetName))
                {
                    sheetobj.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, sheetobj, null);
                    Marshal.ReleaseComObject(sheetobj);
                    break;
                }
                else
                    Marshal.ReleaseComObject(sheetobj);
            }
        }

        public void IgnoreErrorFormat(int rowindex, int startcol, int endcol, int errorformat)
        {
            for (int j = startcol; j < endcol + 1; j++)
            {
                string args = GetExcelColumnName(j) + rowindex.ToString();
                object range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet, new object[] { args });
                object errors = range.GetType().InvokeMember("Errors", BindingFlags.GetProperty, null, range, null);
                object error = errors.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, errors, new object[] { errorformat });  //XlErrorChecks.xlNumberAsText
                error.GetType().InvokeMember("Ignore", BindingFlags.SetProperty, null, error, new object[] { true });

                Marshal.ReleaseComObject(error);
                Marshal.ReleaseComObject(errors);
                Marshal.ReleaseComObject(range);
            }
        }

        public void InsertImage(string filename, double left, double top)
        {
            Image temp = Image.FromFile(filename);
            Single width = temp.Width;
            Single height = temp.Height;
            temp = null;
            if (excelversion >= 2010)
            {
                Single sleft = (Single)left;
                Single stop = (Single)top;
                Single swidth = (Single)width;
                Single sheight = (Single)height;
                object shapes = WorkSheet.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, WorkSheet, null);
                shapes.GetType().InvokeMember("AddPicture", BindingFlags.InvokeMethod, null, shapes, new object[] { filename, 0, -1, sleft, stop, swidth, sheight });
                Marshal.ReleaseComObject(shapes);
            }
            else
            {
                object pictures = WorkSheet.GetType().InvokeMember("Pictures", BindingFlags.GetProperty, null, WorkSheet, null);
                object pic = pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, pictures, new object[] { filename });

                pic.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, pic, new object[] { top });
                pic.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, pic, new object[] { left });

                Marshal.ReleaseComObject(pic);
                Marshal.ReleaseComObject(pictures);
            }

        }

        public void InsertImage(string filename, double left, double top, int width, int height)
        {
            if (excelversion >= 2010)
            {
                Single sleft = (Single)left;
                Single stop = (Single)top;
                Single swidth = (Single)width;
                Single sheight = (Single)height;
                object shapes = WorkSheet.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, WorkSheet, null);
                shapes.GetType().InvokeMember("AddPicture", BindingFlags.InvokeMethod, null, shapes, new object[] { filename, 0, -1, sleft, stop, swidth, sheight });
                Marshal.ReleaseComObject(shapes);
            }
            else
            {
                object pictures = WorkSheet.GetType().InvokeMember("Pictures", BindingFlags.GetProperty, null, WorkSheet, null);
                object pic = pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, pictures, new object[] { filename });

                pic.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, pic, new object[] { top });
                pic.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, pic, new object[] { left });

                Marshal.ReleaseComObject(pic);
                Marshal.ReleaseComObject(pictures);
            }
        }

        //public void InsertImage(int item, string filename, double left, double top, int width, int height)
        //{
        //    if (excelversion >= 2010)
        //    {
        //        Single sleft = (Single)left;
        //        Single stop = (Single)top;
        //        Single swidth = (Single)width;
        //        Single sheight = (Single)height;
        //        object shapes = WorkSheet.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, WorkSheet, null);
        //        shapes.GetType().InvokeMember("AddPicture", BindingFlags.InvokeMethod, null, shapes, new object[] { filename, 0, -1, sleft, stop, swidth, sheight });
        //        Marshal.ReleaseComObject(shapes);
        //    }
        //    else
        //    {
        //        object pictures = WorkSheet.GetType().InvokeMember("Pictures", BindingFlags.GetProperty, null, WorkSheet, null);
        //        object pic = pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, pictures, new object[] { filename });

        //        pic.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, pic, new object[] { top });
        //        pic.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, pic, new object[] { left });

        //        Marshal.ReleaseComObject(pic);
        //        Marshal.ReleaseComObject(pictures);
        //    }

        //    // See if it has worked if PItem is an integer and it's value is not 0 it has worked...
        //    //object PItem = pic.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, pic, null);
        //    //if ((int)PItem > 0)
        //    //{
        //    //Assign PItem again, now to the Item of the Pictures - Collection we just inserted...
        //    //object PItem = pic.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, pic, new object[] { 1 });
        //    // GetShape - Range of Image, now we can move the image to wherever we want to...
        //    //object SRange = pic.GetType().InvokeMember("ShapeRange", BindingFlags.GetProperty, null, pic, null);

        //    //width = int.Parse(pic.GetType().InvokeMember("Width", BindingFlags.GetProperty, null, pic, null).ToString());
        //    //height = int.Parse(pic.GetType().InvokeMember("Height", BindingFlags.GetProperty, null, pic, null).ToString());

        //    // Resize the Shape - Range to 100x100
        //    //SRange.GetType().InvokeMember("Width", BindingFlags.SetProperty, null, SRange, new object[] { width });
        //    //SRange.GetType().InvokeMember("Height", BindingFlags.SetProperty, null, SRange, new object[] { height });
        //    // place the image to the left top corner
        //    //SRange.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, SRange, new object[] { top });
        //    //SRange.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, SRange, new object[] { left });
        //    //}
        //}

        //public enum XlPasteType : int
        //{
        //    xlPasteValues = -4163,
        //    xlPasteFormats = -4122,
        //    xlPasteAll = -4104,
        //}


        public void CopyRows(int rowcount, int startcol, int endcol, int sourcerow, XlPasteType type)
        {
            string sourceAddress = GetExcelColumnName(startcol) + sourcerow.ToString() + ":" + GetExcelColumnName(endcol) + sourcerow.ToString();
            object sourceRange = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet, new object[] { sourceAddress });

            for (int i = sourcerow + 1; i < sourcerow + rowcount; i++)
            {
                string destAddress = GetExcelColumnName(startcol) + i.ToString() + ":" + GetExcelColumnName(endcol) + i.ToString();
                object destRange = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, WorkSheet, new object[] { destAddress });
                CopyRange(sourceRange, destRange, type);

                Marshal.ReleaseComObject(destRange);
            }
            Marshal.ReleaseComObject(sourceRange);
        }

        public void CopyRange(object range_from, object range_to, XlPasteType pasteType)
        {
            range_from.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, range_from, null);
            object[] args = new object[4];
            args[0] = (int)pasteType;
            args[1] = -4142;//Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone
            args[2] = false;
            args[3] = false;
            range_to.GetType().InvokeMember("PasteSpecial", BindingFlags.InvokeMethod, null, range_to, args);
        }

        public void DeleteRow(int index)
        {
            object row = GetRow(index);
            object entireRow = row.GetType().InvokeMember("EntireRow", BindingFlags.GetProperty, null, row, null);
            entireRow.GetType().InvokeMember("Delete", BindingFlags.SetProperty, null, entireRow, new object[] { index });
        }

        public void sheetCopy(string srcname, string destname)
        {
            int count = GetSheetCount();
            for (int i = 1; i <= count; i++)
            {
                object sheetobj = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { i });
                string sheetName = sheetobj.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sheetobj, null).ToString();
                if (srcname.Equals(sheetName))
                {
                    object[] parameters = new object[] { i };
                    object newsheet = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, parameters);

                    parameters = new object[] { Missing.Value, sheetobj };

                    newsheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, newsheet, parameters);
                    Marshal.ReleaseComObject(newsheet);
                    newsheet = GetActiveWorkSheet();
                    newsheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, newsheet, new object[] { destname });
                    object sheetAfter = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { count + 1 });
                    newsheet.GetType().InvokeMember("Move", BindingFlags.InvokeMethod, null, newsheet, new object[] { Missing.Value, sheetAfter });

                    Marshal.ReleaseComObject(sheetAfter);
                    Marshal.ReleaseComObject(sheetobj);
                    Marshal.ReleaseComObject(newsheet);
                    break;
                }
                Marshal.ReleaseComObject(sheetobj);
            }
        }

        public void SetBackgroundPicture(string imagepath)
        {
            object pageSetup = WorkSheet.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, WorkSheet, null);
            object centerHeaderPicture = pageSetup.GetType().InvokeMember("CenterHeaderPicture", BindingFlags.GetProperty, null, pageSetup, null);
            centerHeaderPicture.GetType().InvokeMember("FileName", BindingFlags.SetProperty, null, centerHeaderPicture, new object[] { imagepath });
            WorkSheet.GetType().InvokeMember("SetBackgroundPicture", BindingFlags.Public | BindingFlags.InvokeMethod, null, WorkSheet, new object[] { imagepath });

            Marshal.ReleaseComObject(centerHeaderPicture);
            Marshal.ReleaseComObject(pageSetup);
        }

        public void RearrangeSheet(string basestr, bool isToFirst)
        {
            int realIndex = 1;
            int totcount = GetSheetCount();
            int sheetendindex = totcount;
            for (int idx = 1; idx <= totcount; idx++)
            {
                SetActiveWorkSheet(idx);
                string Name = GetSheetName(idx);
                if (Name.Contains(basestr))
                {
                    object sheetAfter = null;
                    if (isToFirst)
                        sheetAfter = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { realIndex });
                    else
                        sheetAfter = WorkSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, WorkSheets, new object[] { totcount });

                    if (isToFirst)
                        WorkSheet.GetType().InvokeMember("Move", BindingFlags.InvokeMethod, null, WorkSheet, new object[] { Missing.Value, sheetAfter });
                    else
                        WorkSheet.GetType().InvokeMember("Move", BindingFlags.InvokeMethod, null, WorkSheet, new object[] { Missing.Value, sheetAfter });
                    realIndex++;

                    if (isToFirst == false)
                    {
                        idx--;
                        sheetendindex--;
                    }

                    Marshal.ReleaseComObject(sheetAfter);

                    if (idx > sheetendindex)
                        break;
                }
            }
        }


        #region private Wrappers for chart

        private void SetProperty(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            obj.GetType().InvokeMember(sProperty, BindingFlags.SetProperty, null, obj, oParam);
        }
        private object GetProperty(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
        }
        private object GetProperty(object obj, string sProperty, object oValue1, object oValue2)
        {
            object[] oParam = new object[2];
            oParam[0] = oValue1;
            oParam[1] = oValue2;
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
        }
        private object GetProperty(object obj, string sProperty)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, null);
        }
        private object InvokeMethod(object obj, string sProperty, object[] oParam)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
        }
        private object InvokeMethod(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
        }

        public void SetCell(int row, int col, object Value)
        {
            string range = GetRangeStr(row, col);
            object oRange = GetProperty(WorkSheet, "Range", range, Missing.Value);
            SetProperty(oRange, "Value", Value);
        }





        // 요약:
        //     Specifies the chart type.
        public enum ExlChartType
        {
            // 요약:
            //     Scatter
            xlXYScatter = -4169,
            //
            // 요약:
            //     Radar
            xlRadar = -4151,
            //
            // 요약:
            //     Doughnut
            xlDoughnut = -4120,
            //
            // 요약:
            //     3D Pie
            xl3DPie = -4102,
            //
            // 요약:
            //     3D Line
            xl3DLine = -4101,
            //
            // 요약:
            //     3D Column
            xl3DColumn = -4100,
            //
            // 요약:
            //     3D Area
            xl3DArea = -4098,
            //
            // 요약:
            //     Area
            xlArea = 1,
            //
            // 요약:
            //     Line
            xlLine = 4,
            //
            // 요약:
            //     Pie
            xlPie = 5,
            //
            // 요약:
            //     Bubble
            xlBubble = 15,
            //
            // 요약:
            //     Clustered Column
            xlColumnClustered = 51,
            //
            // 요약:
            //     Stacked Column
            xlColumnStacked = 52,
            //
            // 요약:
            //     100% Stacked Column
            xlColumnStacked100 = 53,
            //
            // 요약:
            //     3D Clustered Column
            xl3DColumnClustered = 54,
            //
            // 요약:
            //     3D Stacked Column
            xl3DColumnStacked = 55,
            //
            // 요약:
            //     3D 100% Stacked Column
            xl3DColumnStacked100 = 56,
            //
            // 요약:
            //     Clustered Bar
            xlBarClustered = 57,
            //
            // 요약:
            //     Stacked Bar
            xlBarStacked = 58,
            //
            // 요약:
            //     100% Stacked Bar
            xlBarStacked100 = 59,
            //
            // 요약:
            //     3D Clustered Bar
            xl3DBarClustered = 60,
            //
            // 요약:
            //     3D Stacked Bar
            xl3DBarStacked = 61,
            //
            // 요약:
            //     3D 100% Stacked Bar
            xl3DBarStacked100 = 62,
            //
            // 요약:
            //     Stacked Line
            xlLineStacked = 63,
            //
            // 요약:
            //     100% Stacked Line
            xlLineStacked100 = 64,
            //
            // 요약:
            //     Line with Markers
            xlLineMarkers = 65,
            //
            // 요약:
            //     Stacked Line with Markers
            xlLineMarkersStacked = 66,
            //
            // 요약:
            //     100% Stacked Line with Markers
            xlLineMarkersStacked100 = 67,
            //
            // 요약:
            //     Pie of Pie
            xlPieOfPie = 68,
            //
            // 요약:
            //     Exploded Pie
            xlPieExploded = 69,
            //
            // 요약:
            //     Exploded 3D Pie
            xl3DPieExploded = 70,
            //
            // 요약:
            //     Bar of Pie
            xlBarOfPie = 71,
            //
            // 요약:
            //     Scatter with Smoothed Lines
            xlXYScatterSmooth = 72,
            //
            // 요약:
            //     Scatter with Smoothed Lines and No Data Markers
            xlXYScatterSmoothNoMarkers = 73,
            //
            // 요약:
            //     Scatter with Lines.
            xlXYScatterLines = 74,
            //
            // 요약:
            //     Scatter with Lines and No Data Markers
            xlXYScatterLinesNoMarkers = 75,
            //
            // 요약:
            //     Stacked Area
            xlAreaStacked = 76,
            //
            // 요약:
            //     100% Stacked Area
            xlAreaStacked100 = 77,
            //
            // 요약:
            //     3D Stacked Area
            xl3DAreaStacked = 78,
            //
            // 요약:
            //     100% Stacked Area
            xl3DAreaStacked100 = 79,
            //
            // 요약:
            //     Exploded Doughnut
            xlDoughnutExploded = 80,
            //
            // 요약:
            //     Radar with Data Markers
            xlRadarMarkers = 81,
            //
            // 요약:
            //     Filled Radar
            xlRadarFilled = 82,
            //
            // 요약:
            //     3D Surface
            xlSurface = 83,
            //
            // 요약:
            //     3D Surface (wireframe)
            xlSurfaceWireframe = 84,
            //
            // 요약:
            //     Surface (Top View)
            xlSurfaceTopView = 85,
            //
            // 요약:
            //     Surface (Top View wireframe)
            xlSurfaceTopViewWireframe = 86,
            //
            // 요약:
            //     Bubble with 3D effects
            xlBubble3DEffect = 87,
            //
            // 요약:
            //     High-Low-Close
            xlStockHLC = 88,
            //
            // 요약:
            //     Open-High-Low-Close
            xlStockOHLC = 89,
            //
            // 요약:
            //     Volume-High-Low-Close
            xlStockVHLC = 90,
            //
            // 요약:
            //     Volume-Open-High-Low-Close
            xlStockVOHLC = 91,
            //
            // 요약:
            //     Clustered Cone Column
            xlCylinderColClustered = 92,
            //
            // 요약:
            //     Stacked Cone Column
            xlCylinderColStacked = 93,
            //
            // 요약:
            //     100% Stacked Cylinder Column
            xlCylinderColStacked100 = 94,
            //
            // 요약:
            //     Clustered Cylinder Bar
            xlCylinderBarClustered = 95,
            //
            // 요약:
            //     Stacked Cylinder Bar
            xlCylinderBarStacked = 96,
            //
            // 요약:
            //     100% Stacked Cylinder Bar
            xlCylinderBarStacked100 = 97,
            //
            // 요약:
            //     3D Cylinder Column
            xlCylinderCol = 98,
            //
            // 요약:
            //     Clustered Cone Column
            xlConeColClustered = 99,
            //
            // 요약:
            //     Stacked Cone Column
            xlConeColStacked = 100,
            //
            // 요약:
            //     100% Stacked Cone Column
            xlConeColStacked100 = 101,
            //
            // 요약:
            //     Clustered Cone Bar
            xlConeBarClustered = 102,
            //
            // 요약:
            //     Stacked Cone Bar
            xlConeBarStacked = 103,
            //
            // 요약:
            //     100% Stacked Cone Bar
            xlConeBarStacked100 = 104,
            //
            // 요약:
            //     3D Cone Column
            xlConeCol = 105,
            //
            // 요약:
            //     Clustered Pyramid Column
            xlPyramidColClustered = 106,
            //
            // 요약:
            //     Stacked Pyramid Column
            xlPyramidColStacked = 107,
            //
            // 요약:
            //     100% Stacked Pyramid Column
            xlPyramidColStacked100 = 108,
            //
            // 요약:
            //     Clustered Pyramid Bar
            xlPyramidBarClustered = 109,
            //
            // 요약:
            //     Stacked Pyramid Bar
            xlPyramidBarStacked = 110,
            //
            // 요약:
            //     100% Stacked Pyramid Bar
            xlPyramidBarStacked100 = 111,
            //
            // 요약:
            //     3D Pyramid Column
            xlPyramidCol = 112,
        }

        public void CreateChart(int srow, int scol, int erow, int ecol, 
                                int top, int left, int width, int height, 
                                ExlChartType type)
        {
            object oCharts = InvokeMethod(WorkSheet, "ChartObjects", Missing.Value);
            object oChart = InvokeMethod(oCharts, "Add", new object[] {left, top, width, height } );
            oChart = GetProperty(oChart, "Chart");
            string s_range = GetRangeStr(srow, scol);
            string e_range = GetRangeStr(erow, ecol);
            string range = s_range + ":" + e_range;
            object oRange = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                                                            null, WorkSheet, new object[] { range });
            InvokeMethod(oChart, "SetSourceData", oRange);
            SetProperty(oChart, "ChartType", type);

            
        }

        //public void ChartMiniMax(double Xmin, double Xmax, double Ymin, double Ymax)
        //{
        //    SetProperty(oAxisX, "MinimumScale", Xmin);
        //    SetProperty(oAxisX, "MaximumScale", Xmax);
        //    SetProperty(oAxisY, "MinimumScale", Ymin);
        //    SetProperty(oAxisY, "MaximumScale", Ymax);
        //}


        #endregion


        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        #region Shapes

        public void InsertShapes(ShapeDefine.MsoShapeType type, float left, float top, float width, float height)
        {
            Single sleft = (Single)left;
            Single stop = (Single)top;
            Single swidth = (Single)width;
            Single sheight = (Single)height;
            object shapes = WorkSheet.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, WorkSheet, null);
            shapes.GetType().InvokeMember("AddShape", BindingFlags.InvokeMethod, null, shapes, new object[] {type, sleft, stop, swidth, sheight });
            Marshal.ReleaseComObject(shapes);
        }
                
        #endregion



    }



    /////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// application is busy Error Fixed.
    /// https://msdn.microsoft.com/en-us/library/ms228772.aspx
    /// </summary>

    public class MessageFilter : IOleMessageFilter
    {
        //
        // Class containing the IOleMessageFilter
        // thread error-handling functions.

        // Start the filter.
        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, out oldFilter);
            Thread.Sleep(3000);
        }

        // Done with the filter, close it.
        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        //
        // IOleMessageFilter functions.
        // Handle incoming thread requests.
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
          lpInterfaceInfo)
        {
            //Return the flag SERVERCALL_ISHANDLED.
            //return 0;
            return 1;
        }

        // Thread call was rejected, so try again.
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr
          hTaskCallee, int dwTickCount, int dwRejectType)
        {
            //if (dwRejectType == 2)
            //// flag = SERVERCALL_RETRYLATER.
            //{
            //    // Retry the thread call immediately if return >=0 & 
            //    // <100.
            //    return 99;
            //}
            //// Too busy; cancel call.
            //return -1;

            int retVal = -1;
            Debug.WriteLine("RetryRejectedCall");
            if (MessageBox.Show("Office operation was rejected. Close any Dialog Box and try again.", "Alert", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                retVal = 1;
            }
            return retVal;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
          int dwTickCount, int dwPendingType)
        {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            //return 2;
            Debug.WriteLine("MessagePending");
            return 1;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int
          CoRegisterMessageFilter(IOleMessageFilter newFilter, out 
          IOleMessageFilter oldFilter);
    }

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }
}
