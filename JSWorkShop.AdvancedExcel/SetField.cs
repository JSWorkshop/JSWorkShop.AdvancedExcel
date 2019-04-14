using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace JSWorkShop.AdvancedExcel
{
    public class SetField : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<int> RowIndex { get; set; }

        [Category("Input")]
        public InArgument<string> ColLetter { get; set; }

        [Category("Input")]
        public InArgument<int> ColIndex { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FieldValue { get; set; }


        [Category("Output")]
        public OutArgument<Boolean> FieldSet { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Load Input Fields unto local variables
            string fieldValue = FieldValue.Get(context);
            string filePath = FilePath.Get(context);

            string colLetter = ColLetter.Get(context);
            int colIndex = ColIndex.Get(context);
            int rowIndex = RowIndex.Get(context);

            // Initialize Tools create for this project (Class Tools or Tools.cs)
            Tools toolKit = new Tools();

            //Validate if Column Letter is empty and Column Index is 0 to return an error or continue
            if (String.IsNullOrEmpty(colLetter) && colIndex == 0)
            {
                FieldSet.Set(context, false);
                throw new ArgumentNullException();
            }
            else {
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

                        //Validate if Column Letter is empty to determine if to use the column letter or the column index input fields
                        if (String.IsNullOrEmpty(colLetter))
                        {
                            //Set Field using Column Index and Row Index
                            ((Excel.Range)xlWorksheet.Cells[rowIndex, colIndex]).Value2 = fieldValue;
                        }
                        else
                        {
                            //Set Field using Column Letter and Row Index
                            ((Excel.Range)xlWorksheet.Cells[rowIndex, toolKit.GetColNumber(colLetter)]).Value2 = fieldValue;
                        }

                        //Return that Field was Set
                        FieldSet.Set(context, true);

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
                        FieldSet.Set(context, false);
                        throw;
                    }
                }
                else
                {
                    FieldSet.Set(context, false);
                    throw new ArgumentNullException();
                }
            }
        }
    }
}

