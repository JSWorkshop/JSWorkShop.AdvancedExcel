using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace JSWorkShop.AdvancedExcel
{
    public class SetHeader : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> HeaderName { get; set; }

        [Category("Input")]
        public InArgument<string> ColLetter { get; set; }

        [Category("Input")]
        public InArgument<int> ColIndex { get; set; }

        [Category("Output")]
        public OutArgument<Boolean> HeaderSet { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string headerName = HeaderName.Get(context);
            string filePath = FilePath.Get(context);

            string colLetter = ColLetter.Get(context);
            int colIndex = ColIndex.Get(context);

            // Initialize Tools create for this project (Class Tools or Tools.cs)
            Tools toolKit = new Tools();

            //Validate if Column Letter is empty and Column Index is 0 to return an error or continue
            if (String.IsNullOrEmpty(colLetter) && colIndex == 0)
            {
                //We require at least one of the two Column values to process, if not return false and throw an exception
                HeaderSet.Set(context, false);
                throw new ArgumentNullException();
            }
            else
            {
                //If one exists continue

                //Validate that file path is not empty
                if (!String.IsNullOrEmpty(filePath))
                {
                    try
                    {
                        //Initialize Excel Interop objects;
                        object m = Type.Missing;
                        Application xlApp = new Application();
                        Workbooks xlWorkbooks = xlApp.Workbooks;
                        _Workbook xWB;

                        //check if file Exists if not Create the file first
                        if (!File.Exists(filePath))
                        {
                            xWB = xlWorkbooks.Add(XlWBATemplate.xlWBATWorksheet);
                            xWB.SaveAs(filePath, XlFileFormat.xlOpenXMLWorkbook, m, m, false, false, XlSaveAsAccessMode.xlShared, false, false, m, m, m);
                            xWB.Close();
                        }
                        
                        Workbook xlWorkbook = xlWorkbooks.Open(filePath, m, false, m, m, m, m, m, m, m, m, m, m, m, m);
                        _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        Range xlRange = xlWorksheet.UsedRange;
                            

                        //Validate if Column Letter is empty to determine if to use the column letter or the column index input fields
                        if (String.IsNullOrEmpty(colLetter))
                        {
                            //Set Header using Column Index
                            ((Excel.Range)xlWorksheet.Cells[1, colIndex]).Value2 = headerName;
                        }
                        else
                        {
                            //Set Header using Column Letter
                            ((Excel.Range)xlWorksheet.Cells[1, toolKit.GetColNumber(colLetter)]).Value2 = headerName;
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

                        HeaderSet.Set(context, true);

                    }
                    catch
                    {
                        HeaderSet.Set(context, false);
                        throw;
                    }
                }
                else
                {
                    HeaderSet.Set(context, false);
                    throw new ArgumentNullException();
                }
            }
        }
    }
}

