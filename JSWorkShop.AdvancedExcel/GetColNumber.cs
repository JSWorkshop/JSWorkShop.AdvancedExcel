using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;

namespace JSWorkShop.AdvancedExcel
{
    public class GetColNumber : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ColLetter { get; set; }

        [Category("Output")]
        public OutArgument<int> ColIndex { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Initialize Tools create for this project (Class Tools or Tools.cs)
            Tools toolKit = new Tools();

            // Return the Index(Numeric/Int) representation of a Column Letter
            ColIndex.Set(context, toolKit.GetColNumber(ColLetter.Get(context)));
        }
    }
}

