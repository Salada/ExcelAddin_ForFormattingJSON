using JssonPortAddIn.Model.ExcelModel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Tools.Excel;

namespace JssonPortAddIn.Model
{
    class ExcelFileModel : IConvertableJson
    {
        private List<ExcelWorkSheetModel> list;
        private dynamic baseInfo;
        public string FileName
        {
            get;
            set;
        }

        public ExcelFileModel()
        {
            InitializeWorkBookBaseInfo();
            
            InitializeObject();
        }

        public ExcelFileModel(JObject jsoneObject)
        {
            InitializeWorkBookBaseInfo();
            this.FileName = jsoneObject.Property("FileName").Value.ToString();

            // TODO: i think more specify requirement when workbook is already opened concerned jsona format.
            dynamic newWorkBook = this.baseInfo.Workbooks.Add();
            newWorkBook.Activate();
            
            foreach(JObject jWorksheet in jsoneObject.Property("WorkBook").Value)
            {
                dynamic newWorkSheet = this.baseInfo.WorkSheets.Add();
                list.Add(new ExcelWorkSheetModel(jWorksheet, newWorkSheet));
            }
        }

        private void InitializeWorkBookBaseInfo()
        {
            list = new List<ExcelWorkSheetModel>();
            this.baseInfo = Globals.ThisAddIn.Application;
        }
        private void InitializeObject()
        {
            dynamic worksheets = this.baseInfo.WorkSheets;
            for (int i = 0; i < this.baseInfo.WorkSheets.Count; ++i)
            {
                dynamic worksheet = worksheets[i + 1];
                list.Add(new ExcelWorkSheetModel(worksheet));
            }
        }

        public Newtonsoft.Json.Linq.JToken ToJsonObject()
        {
            return this.ToJsonObject(this.FileName ?? "ExcelJsonFormatInfo");
        }

        public Newtonsoft.Json.Linq.JToken ToJsonObject(string fileName)
        {
            FileName = fileName;
            return new JObject(
                new JProperty("FileName", fileName),
                new JProperty("WorkBook",
                    new JArray(
                        from o in list
                        select o.ToJsonObject()
                        )
                    )
                );
        }
    }
}
