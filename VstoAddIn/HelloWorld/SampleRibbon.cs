using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HelloWorld
{
    public partial class SampleRibbon
    {
        private void SampleRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnHello_Click(object sender, RibbonControlEventArgs e)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Range["A1"].Cells.Value = "Hello World!";
        }
    }
}
