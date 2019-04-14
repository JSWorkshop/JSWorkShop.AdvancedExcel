using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities;
using System.IO;
using System.ComponentModel;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSWorkShop.AdvancedExcel
{
    public class BlankTemplate : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InTemp { get; set; }

        [Category("Output")]
        public OutArgument<Boolean> OutTemp { get; set; }

        protected override void Execute(CodeActivityContext context)
        {}
    }
}

