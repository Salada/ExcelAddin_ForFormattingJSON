using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Dynamic;
using System.Collections;
using JssonPortAddIn.Model;

namespace JssonPortAddIn
{
    public partial class RbnFileManage
    {
        private void RbnFileManage_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            var excel = new ExcelFileModel();
            MessageBox.Show(excel.ToJsonObject().ToString());
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
