using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace JSWorkShop.AdvancedExcel
{
    public class GetRowStatus : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        public InArgument<string> ColLetter { get; set; }

        [Category("Input")]
        public InArgument<int> ColIndex { get; set; }

        [Category("Input")]
        public InArgument<int> RowIndex { get; set; }

        [Category("Output")]
        public OutArgument<string> Status { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string filePath = FilePath.Get(context);
            string colLetter = ColLetter.Get(context);
            int colIndex = ColIndex.Get(context);
            int rowIndex = RowIndex.Get(context);

            //Validate that file path is not empty
            if (!String.IsNullOrEmpty(filePath))
            {
                //Validate that file exists
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

                        // Initialize Tools create for this project (Class Tools or Tools.cs)
                        Tools toolKit = new Tools();

                        //Validate if Column Letter is empty and Column Index is 0 to return an error or continue
                        if (String.IsNullOrEmpty(colLetter) && colIndex == 0)
                        {
                            //We require at least one of the two Column values to process, if not return false and throw an exception
                            Status.Set(context, "Error: No Column Letter or Index provided.");
                            throw new ArgumentNullException();
                        }
                        else
                        {
                            //Validate if Column Letter is empty to determine if to use the column letter or the column index input fields
                            if (String.IsNullOrEmpty(colLetter))
                            {
                                //Get Status Value using Column Index
                                Status.Set(context, ((Excel.Range)xlWorksheet.Cells[rowIndex, colIndex]).Value);
                            }
                            else
                            {
                                //Get Status Value using Column Letter(s)
                                Status.Set(context, ((Excel.Range)xlWorksheet.Cells[rowIndex, toolKit.GetColNumber(colLetter)]).Value);
                            }
                        }

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

