using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Dynamic;
using System.Collections;

namespace JssonPortAddIn
{
    public partial class RbnFileManage
    {
        private void RbnFileManage_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            
            dynamic worksheets = Globals.ThisAddIn.Application.Worksheets;

            dynamic a = worksheets[2];

            var table = new ExcelTableClass(worksheets[2]);
            MessageBox.Show("Hello");
        }
    }
}
