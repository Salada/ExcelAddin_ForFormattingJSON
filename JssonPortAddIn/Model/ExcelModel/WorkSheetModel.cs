using JssonPortAddIn.Model.ExcelModel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JssonPortAddIn
{
    class WorkSheetModel : IConvertableJson
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

        public WorkSheetModel()
        {
            list = new List<ExcelTableModel>();
        }

        public WorkSheetModel(dynamic worksheet)
            : this()
        {
            this.WorkSheet = worksheet;
            this.SheetName = this.WorkSheet.Name;
            this.CodeName = this.WorkSheet.CodeName;

            InitializeObject();
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

        public Newtonsoft.Json.Linq.JObject ToJsonObject()
        {
            return new JObject(
                new JProperty("SheetName", this.SheetName),
                new JProperty("CodeName", this.CodeName),
                new JProperty("TableList",
                    new JArray(
                        from o in list

                        select o.ToJsonObject()
                        )
                    )
                );
        }
    }
}
