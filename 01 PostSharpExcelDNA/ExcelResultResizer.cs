namespace PostSharpExcelDNA
{
    using ExcelDna.Integration;

    using System;
    using System.Collections.Generic;

    internal class ExcelResultResizer : XlCall
    {
        private static readonly HashSet<ExcelReference> ExcelReferenceBeingProcessed = new HashSet<ExcelReference>();

        internal static void ResizeAndAutoFormat(object[,] array)
        {
            var caller = Excel(xlfCaller) as ExcelReference;
            if (caller == null)
            {
                return;
            }

            int rows = array.GetLength(0);
            int columns = array.GetLength(1);

            if (rows == 0 || columns == 0)
            {
                return;
            }

            if ((caller.RowLast - caller.RowFirst + 1 == rows) &&
                (caller.ColumnLast - caller.ColumnFirst + 1 == columns))
            {
                // Size is already OK - just return result
                return;
            }

            var rowLast = caller.RowFirst + rows - 1;
            var columnLast = caller.ColumnFirst + columns - 1;

            // Check for the sheet limits
            if (rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 || columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1)
            {
                // Can't resize - goes beyond the end of the sheet - just return #VALUE
                return;
            }

            // Create a reference of the right size
            var target = new ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId);

            lock (ExcelReferenceBeingProcessed)
            {
                if (ExcelReferenceBeingProcessed.Contains(target))
                {
                    return;
                }

                ExcelReferenceBeingProcessed.Add(target);
            }

            ExcelAsyncUtil.QueueAsMacro(
                () =>
                {
                    DoResize(target); // Will trigger a recalc by writing formula

                    lock (ExcelReferenceBeingProcessed)
                    {
                        ExcelReferenceBeingProcessed.Remove(target);
                    }
                });
        }

        private static void DoResize(ExcelReference target)
        {
            try
            {
                using (new ExcelEchoOffHelper())
                {
                    using (new ExcelCalculationManualHelper())
                    {
                        var firstCell = new ExcelReference(
                            target.RowFirst,
                            target.RowFirst,
                            target.ColumnFirst,
                            target.ColumnFirst,
                            target.SheetId);

                        // Get the formula in the first cell of the target
                        var formula = (string)Excel(XlCall.xlfGetCell, 41, firstCell);
                        var isFormulaArray = (bool)Excel(XlCall.xlfGetCell, 49, firstCell);
                        if (isFormulaArray)
                        {
                            // Select the sheet and firstCell - needed because we want to use SelectSpecial.
                            using (new ExcelSelectionHelper(firstCell))
                            {
                                // Extend the selection to the whole array and clear
                                XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                                var oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

                                oldArray.SetValue(ExcelEmpty.Value);
                            }
                        }

                        // Get the formula and convert to R1C1 mode
                        var isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
                        string formulaR1C1 = formula;
                        if (!isR1C1Mode)
                        {
                            // Set the formula into the whole target
                            formulaR1C1 =
                                (string)
                                Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, firstCell);
                        }

                        // Must be R1C1-style references
                        object ignoredResult;

                        var retval = TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);

                        if (retval != XlReturn.XlReturnSuccess)
                        {
                            var firstCellAddress = (string)Excel(XlCall.xlfReftext, firstCell, true);
                            XlCall.Excel(
                                XlCall.xlcAlert,
                                string.Format(
                                    "Cannot resize array formula at {0} - result might overlap another array.",
                                    firstCellAddress));

                            firstCell.SetValue(string.Format("'{0}", formula));
                        }
                    }
                }
            }
            finally
            {
            }
        }

        // RIIA-style helpers to deal with Excel selections    
        // Don't use if you agree with Eric Lippert here: http://stackoverflow.com/a/1757344/44264
        private class ExcelEchoOffHelper : XlCall, IDisposable
        {
            private readonly object oldEcho;

            public ExcelEchoOffHelper()
            {
                this.oldEcho = XlCall.Excel(XlCall.xlfGetWorkspace, 40);
                XlCall.Excel(XlCall.xlcEcho, false);
            }

            public void Dispose()
            {
                XlCall.Excel(XlCall.xlcEcho, this.oldEcho);
            }
        }

        private class ExcelCalculationManualHelper : XlCall, IDisposable
        {
            private readonly object oldCalculationMode;

            public ExcelCalculationManualHelper()
            {
                this.oldCalculationMode = XlCall.Excel(XlCall.xlfGetDocument, 14);
                XlCall.Excel(XlCall.xlcOptionsCalculation, 3);
            }

            public void Dispose()
            {
                XlCall.Excel(XlCall.xlcOptionsCalculation, this.oldCalculationMode);
            }
        }

        // Select an ExcelReference (perhaps on another sheet) allowing changes to be made there.
        // On clean-up, resets all the selections and the active sheet.
        // Should not be used if the work you are going to do will switch sheets, amke new sheets etc.
        private class ExcelSelectionHelper : XlCall, IDisposable
        {
            private readonly object oldSelectionOnActiveSheet;
            private readonly object oldActiveCellOnActiveSheet;

            private readonly object oldSelectionOnRefSheet;
            private readonly object oldActiveCellOnRefSheet;

            public ExcelSelectionHelper(ExcelReference refToSelect)
            {
                // Remember old selection state on the active sheet
                this.oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                this.oldActiveCellOnActiveSheet = XlCall.Excel(XlCall.xlfActiveCell);

                // Switch to the sheet we want to select
                var refSheet = (string)XlCall.Excel(XlCall.xlSheetNm, refToSelect);
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { refSheet });

                // record selection and active cell on the sheet we want to select
                this.oldSelectionOnRefSheet = XlCall.Excel(XlCall.xlfSelection);
                this.oldActiveCellOnRefSheet = XlCall.Excel(XlCall.xlfActiveCell);

                // make the selection
                XlCall.Excel(XlCall.xlcFormulaGoto, refToSelect);
            }

            public void Dispose()
            {
                // Reset the selection on the target sheet
                XlCall.Excel(XlCall.xlcSelect, this.oldSelectionOnRefSheet, this.oldActiveCellOnRefSheet);

                // Reset the sheet originally selected
                var oldActiveSheet = (string)XlCall.Excel(XlCall.xlSheetNm, this.oldSelectionOnActiveSheet);
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { oldActiveSheet });

                // Reset the selection in the active sheet (some bugs make this change sometimes too)
                XlCall.Excel(XlCall.xlcSelect, this.oldSelectionOnActiveSheet, this.oldActiveCellOnActiveSheet);
            }
        }
    }
}