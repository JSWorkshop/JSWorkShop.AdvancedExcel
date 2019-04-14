using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;

namespace JSWorkShop.AdvancedExcel
{
    public class GetColLetter: CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<int> ColIndex { get; set; }

        [Category("Output")]
        public OutArgument<string> ColLetter { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Initialize Tools create for this project (Class Tools or Tools.cs)
            Tools toolKit = new Tools();

            // Return the Letter Representation of the ColIndex
            ColLetter.Set(context, toolKit.GetColLetter(ColIndex.Get(context)));
        }
    }
}

