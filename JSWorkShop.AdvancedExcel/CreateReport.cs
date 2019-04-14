using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace JSWorkShop.AdvancedExcel
{
    public class CreateReport : CodeActivity
    {   
        //INPUTS
        //File Path to create report
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }
        //hItem[String] - Name of the column header for the item(usually an account number of other identifier)
        [Category("Input")]
        [RequiredArgument]        
        public InArgument<string> HeaderName { get; set; }
        //Step1...X [String] - Header values for each step, as applicable
        [Category("Input")]
        public InArgument<string[]> StepsName { get; set; }

        //OUTPUTS:
        //NextRow[Int32] - Value of the next row to process.This value is build based on the next available row where the status is not complete.
        [Category("Output")]
        public OutArgument<Int32> NextRow { get; set; }
        //StatusCol[String] - Column letter for the status column.
        [Category("Output")]
        public OutArgument<string> StatusCol { get; set; }
        //out_Step{ 1...X} - Column letter for each step column, in the order it was provided
        [Category("Output")]
        public OutArgument<List<string>> StepsCols { get; set; }
        //False if an error is thrown
        [Category("Output")]
        public OutArgument<Boolean> ReportSet { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string headerName = HeaderName.Get(context);
            string filePath = FilePath.Get(context);
            string[] steps = StepsName.Get(context);
            List<string> stepcols = new List<string>();
            int stepcounter;

            // Initialize Tools create for this project (Class Tools or Tools.cs)
            Tools toolKit = new Tools();

            //Validate if HeaderName Exists
            if (String.IsNullOrEmpty(headerName))
            {
                //We require at least one of the two Column values to process, if not return false and throw an exception
                ReportSet.Set(context, false);
                throw new ArgumentNullException("Missing a Header Name");
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

                        //Create Header and Status Column
                        ((Range)xlWorksheet.Cells[1, 1]).Value2 = headerName;
                        ((Range)xlWorksheet.Cells[1, 2]).Value2 = "Status";

                        //Set Status Column
                        StatusCol.Set(context, "B");

                        #region Calculate Steps Array values
                        int StepsCount;
                        try
                        {
                            StepsCount = steps.GetUpperBound(0);
                        }catch
                        {
                            StepsCount = 0;
                        }
                        #endregion

                        Tools mytools = new Tools();
                        
                        //Process Steps Headers
                        stepcounter = 3;
                        if (StepsCount > 0)
                        {
                            foreach (string step in steps)
                            {
                                //Set Header using Column Index
                                ((Range)xlWorksheet.Cells[1, stepcounter]).Value2 = step;
                                stepcols.Add(mytools.GetColLetter(stepcounter));
                                stepcounter++;
                            }

                            StepsCols.Set(context, stepcols);
                        }

                        //FIND NEXTROW (Next blank row by Column A)
                        for (stepcounter = 2; stepcounter < 1000000; stepcounter++)
                        {
                            if (String.IsNullOrEmpty(((Range)xlWorksheet.Cells[stepcounter, 1]).Value2))
                            {
                                NextRow.Set(context, stepcounter);
                                break;
                            }
                        }

                        if(NextRow.Get(context)==0) {
                            NextRow.Set(context, 2);
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

                        ReportSet.Set(context, true);

                    }
                    catch(Exception ex)
                    {
                        ReportSet.Set(context, false);
                        throw ex;
                    }
                }
                else
                {
                    ReportSet.Set(context, false);
                    throw new ArgumentNullException("Missing FilePath to create file.");
                }
            }
        }
    }
}

