using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Dynamic;
using System.Collections;
using JssonPortAddIn.Model;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JssonPortAddIn
{
    public partial class RbnFileManage
    {
        private string savedFileDirectory;
        private string savedFileName;
        private void RbnFileManage_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            var excel = new ExcelFileModel();

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Custom json format|*.jsone";
            saveFileDialog1.Title = "Save file";
            saveFileDialog1.FileName = this.savedFileName;
            saveFileDialog1.InitialDirectory = this.savedFileDirectory;
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                using(Stream fs = saveFileDialog1.OpenFile())
                using(StreamWriter file = new StreamWriter(fs))
                using(JsonTextWriter writer = new JsonTextWriter(file))
                {
                    if (!saveFileDialog1.FileName.EndsWith(".jsone"))
                        saveFileDialog1.FileName += ".jsone";

                    this.savedFileName = saveFileDialog1.FileName.Substring(saveFileDialog1.FileName.LastIndexOf('\\') + 1);
                    this.savedFileDirectory = saveFileDialog1.InitialDirectory;
                    file.Write(excel.ToJsonObject(this.savedFileName).ToString());
                }
            }

            //MessageBox.Show(excel.ToJsonObject().ToString());
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            JObject jsoneObject;

            openFileDialog1.Filter = "Custom json format|*.jsone";
            openFileDialog1.Title = "Load jsona file";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using(Stream fs = openFileDialog1.OpenFile())
                using(StreamReader file = new StreamReader(fs))
                using(JsonTextReader reader = new JsonTextReader(file))
                {
                    jsoneObject = (JObject)JToken.ReadFrom(reader);
                    var excelFileModel = new ExcelFileModel(jsoneObject);
                }
            }

            
            
        }
    }
}
