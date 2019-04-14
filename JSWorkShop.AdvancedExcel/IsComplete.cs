using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace JSWorkShop.AdvancedExcel
{
    public class IsComplete : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        public OutArgument<Boolean> isComplete { get; set; }

        //DEBUG Value only used when trying to debug code
        [Category("Output")]
        public OutArgument<string> DebugOut { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string filePath = FilePath.Get(context);

            //Validate that file path is not empty
            if (!String.IsNullOrEmpty(filePath))
            {
                if (File.Exists(filePath))
                {
                    try
                    {
                        //Initialize Excel Interop objects;
                        object m = Type.Missing;
                        Excel.Application xlApp = new Excel.Application();
                        Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
                        Excel.Workbook xlWorkbook = xlWorkbooks.Open(filePath, m, true, m, m, m, m, m, m, m, m, m, m, m, m);
                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        Excel.Range xlRange = xlWorksheet.UsedRange;

                        // Gather Row and Col counts
                        Int32 lastRow = xlRange.Rows.Count;
                        Int32 lCol = xlRange.Columns.Count;
                        Int32 sCol = lCol + 1;

                        // Create local variables for processing
                        int i;
                        string vLstCol;

                        //iterate by the existing columns to determine if a current Status column exists
                        for (i = 1; i <= lCol; i++)
                        {
                            vLstCol = xlWorksheet.Cells[1, i].Value;
                            if (vLstCol == "Status")
                            {
                                sCol = i;
                                lCol = i - 1;
                            }
                        }

                        //Count all rows that Status Column has the value of "Complete"
                        int countComplete = (int)xlApp.WorksheetFunction.CountIf(xlRange.Columns[sCol, m], "Complete") + 1;
                        //Count all rows that are not empty on the first column (Item-Index column)
                        int countNotNull = (int)xlApp.WorksheetFunction.CountA(xlRange.Columns[1, m]);

                        //Validate that complete rows match the lastRow count
                        if (countComplete == countNotNull)
                        {
                            //Return True for Complete
                            isComplete.Set(context, true);
                        }
                        else
                        {
                            //Return False for not complete
                            isComplete.Set(context, false);
                        }

                        //DebugOut Output
                        DebugOut.Set(context, "CC: " + countComplete.ToString() + " | CA: " + countNotNull.ToString() + " | LR: " +lastRow.ToString());

                        //Close Workbook no Save
                        xlWorkbook.Close(false, m, m);

                        //CLOSE AND GARBAGE COLLECT
                        Marshal.ReleaseComObject(xlWorksheet);
                        xlWorksheet = null;
                        Marshal.ReleaseComObject(xlWorkbooks);
                        xlWorkbooks = null;
                        Marshal.ReleaseComObject(xlWorkbook);
                        xlWorkbook = null;

                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlApp);
                        xlApp = null;

                        GC.Collect(); //Garbage Collect
                        GC.WaitForPendingFinalizers(); //Wait until Garbage Collect completes

                    }
                    catch
                    {
                        throw;
                    }
                }
                else
                {
                    throw new FileNotFoundException();
                }
            }
            else
            {
                throw new ArgumentNullException();
            }
        }
    }
}

