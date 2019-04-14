using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace JSWorkShop.AdvancedExcel
{
    public class GetBounds : CodeActivity
    {
        //Prerequisites:
        //    Column A must be an index column with non blank cells for every data row.
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        public OutArgument<Int32> LastRow { get; set; }

        [Category("Output")]
        public OutArgument<string> LastCol { get; set; }

        [Category("Output")]
        public OutArgument<string> StatusCol { get; set; }

        [Category("Output")]
        public OutArgument<Int32> NextRow { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string filePath = FilePath.Get(context);

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
                        Excel.Workbook xlWorkbook = xlWorkbooks.Open(filePath, m, false, m, m, m, m, m, m, m, m, m, m, m, m);
                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        //Excel.Range xlRange = xlWorksheet.UsedRange;
                        Excel.Range xlRange = xlWorksheet.get_Range("a1").EntireRow.EntireColumn;

                        xlApp.DisplayAlerts = false;

                        // Initialize Tools create for this project (Class Tools or Tools.cs)
                        Tools toolKit = new Tools();

                        // Gather Row and Col counts
                        Console.WriteLine("Start");

                        Int32 lastRow = (Int32)xlApp.WorksheetFunction.CountA(xlRange.get_Range("a1").EntireColumn);
                        Int32 lCol = xlRange.Columns.Count;
                        Int32 sCol = lCol;

                        Console.WriteLine(lastRow.ToString());

                        // Create local variables for processing
                        int i;
                        string vLstCol;
                        string statusValue;
                        var eValue = "";

                        //iterate by the existing columns to determine if a current Status column exists
                        for (i = 1; i <= lCol; i++)
                        {
                            vLstCol = xlWorksheet.Cells[1, i].Value;
                            if (vLstCol == "Status" || String.IsNullOrEmpty(vLstCol))
                            {
                                if (vLstCol == "Status") {
                                    sCol = i;
                                    lCol = i - 1;
                                }
                                else
                                {
                                    lCol = i;
                                    sCol = i + 1;
                                }
                            }
                        }

                        //set Status header (or reset if one already exists)
                        ((Excel.Range)xlWorksheet.Cells[1, sCol]).Value2 = "Status";

                        //create and assign Last Column value
                        String lastCol = toolKit.GetColLetter(lCol);
                        LastCol.Set(context, lastCol);

                        //create and assign Status Column value
                        String statusCol = toolKit.GetColLetter(lCol + 1);
                        StatusCol.Set(context, statusCol);

                        //Assign Last Row Value
                        LastRow.Set(context, lastRow);

                        Console.WriteLine("Set status range");
                        Excel.Range xlStatus = xlRange.get_Range(statusCol + "1").EntireColumn;

                        Console.WriteLine(xlStatus.Rows.Count.ToString() + "/" + xlStatus.Columns.Count.ToString());
                        //Iterate through the rows to determine if status is complete or not (will determine next item to work on)
                        i = 1;

                        Console.WriteLine("Get Next Row");
                        for( i = 1; i < xlStatus.Rows.Count; i++) // while using eachrow on xlStatus seemed the best approach this allow for better value control
                        {
                            eValue = ((Excel.Range)xlWorksheet.Cells[i, sCol]).Value2;
                            //Console.WriteLine(eValue.GetType().ToString());
                            statusValue = "";
                            
                            if (eValue != null) {
                                statusValue = eValue.ToString();
                                statusValue = statusValue.ToLower().Trim();
                            }

                            Console.WriteLine(i.ToString() + "/" + sCol.ToString() + " - " + statusValue);
                            if (statusValue != "complete" && statusValue != "status")
                            {
                                //Assign Next Row value
                                NextRow.Set(context, i);
                                break;
                            }
                        }

                        //Save and Close Workbook
                        xlWorkbook.Save();
                        xlWorkbook.Close(true, m, m);

                        //Close all Excel Interop objects and perform Garbage Collect
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

