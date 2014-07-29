using JssonPortAddIn.Model.ExcelModel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Tools.Excel;

namespace JssonPortAddIn
{
    class ExcelWorkSheetModel : IConvertableJson
    {
        private List<ExcelTableModel> list;
        public dynamic WorkSheet
        {
            get;
            private set;
        }


        public string CodeName
        {
            get;
            set;
        }

        public string SheetName
        {
            get;
            set;
        }

        public ExcelWorkSheetModel()
        {
            list = new List<ExcelTableModel>();
        }

        public ExcelWorkSheetModel(JObject jWorksheet, dynamic worksheet) : this()
        {
            this.SheetName = worksheet.Name = jWorksheet["SheetName"].ToString();
            this.CodeName = jWorksheet["CodeName"].ToString();
            this.WorkSheet = worksheet;

            foreach (JObject obj in jWorksheet["TableList"])
            {
                ExcelTableModel table = new ExcelTableModel(this, obj);
            }
            
            
        }

        public ExcelWorkSheetModel(dynamic worksheet)
            : this()
        {
            InitializeWorkSheetBaseInfo(worksheet);

            InitializeObject();
        }

        private void InitializeWorkSheetBaseInfo(dynamic worksheet)
        {
            this.WorkSheet = worksheet;
            this.SheetName = this.WorkSheet.Name;
            this.CodeName = this.WorkSheet.CodeName;
        }
        private void InitializeObject()
        {
            try
            {   
                var newModel = new ExcelTableModel(this);
                list.Add(newModel);
            }
            catch (System.Exception ex)
            {
            }
        }

        public Newtonsoft.Json.Linq.JToken ToJsonObject()
        {
            return new JObject(
                new JProperty("SheetName", this.SheetName),
                new JProperty("CodeName", this.CodeName),
                new JProperty("TableList",
                        from o in list
                        select o.ToJsonObject()
                    )
                );
        }
    }
}
