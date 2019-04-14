using System;
using System.Activities;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace JSWorkShop.AdvancedExcel
{
    public class SetRowStatus : CodeActivity
    {
        public enum eStatus
        {
            Pending,
            Complete,
            Failed,
            Retried,
            [Display(Name = "N/A" )]
            NA
        }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        public InArgument<string> ColLetter { get; set; }

        [Category("Input")]
        public InArgument<int> ColIndex { get; set; }

        [Category("Input")]
        public InArgument<int> RowIndex { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public eStatus StatusValue { get; set; }

        [Category("Output")]
        public OutArgument<Boolean> StatusSet { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string filePath = FilePath.Get(context);
            string colLetter = ColLetter.Get(context);
            int colIndex = ColIndex.Get(context);
            int rowIndex = RowIndex.Get(context);
            eStatus setStatus = StatusValue;

            //Validate that file path is not empty
            if (!String.IsNullOrEmpty(filePath))
            {
                try
                {
                    //Initialize Excel Interop objects;
                    object m = Type.Missing;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
                    Excel._Workbook xWB;

                    //check if file Exists if not Create the file first
                    if (!File.Exists(filePath))
                    {
                        xWB = xlWorkbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                        xWB.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook, m, m, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, m, m, m);
                        xWB.Close();
                    }

                    Excel.Workbook xlWorkbook = xlWorkbooks.Open(filePath, m, false, m, m, m, m, m, m, m, m, m, m, m, m);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    // Initialize Tools create for this project (Class Tools or Tools.cs)
                    Tools toolKit = new Tools();

                    //Validate if Column Letter is empty and Column Index is 0 to return an error or continue
                    if (String.IsNullOrEmpty(colLetter) && colIndex == 0)
                    {
                        //We require at least one of the two Column values to process, if not return false and throw an exception
                        StatusSet.Set(context, false);
                        throw new ArgumentNullException();
                    }
                    else
                    {
                        //Validate if Column Letter is empty to determine if to use the column letter or the column index input fields
                        if (String.IsNullOrEmpty(colLetter))
                        {
                            //Set Status Value using Column Index
                            ((Excel.Range)xlWorksheet.Cells[rowIndex, colIndex]).Value2 = StatusValue.ToString();
                            StatusSet.Set(context, true);
                        }
                        else
                        {
                            //Set Status Value using Column Letter(s)
                            ((Excel.Range)xlWorksheet.Cells[rowIndex, toolKit.GetColNumber(colLetter)]).Value2 = StatusValue.ToString();
                            StatusSet.Set(context, false);
                        }
                    }

                    //Save and Close Workbook
                    xlWorkbook.Save();
                    xlWorkbook.Close(true, m, m);

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
                throw new ArgumentNullException();
            }
        }
    }
}

