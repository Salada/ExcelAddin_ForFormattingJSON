using JssonPortAddIn.Model.ExcelModel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JssonPortAddIn.Model
{
    class ExcelFileModel : IConvertableJson
    {
        private List<WorkSheetModel> list;
        private dynamic baseInfo;

        public ExcelFileModel()
        {
            list = new List<WorkSheetModel>();
            this.baseInfo = Globals.ThisAddIn.Application;

            InitializeObject();
        }
        
        private void InitializeObject()
        {
            dynamic worksheets = this.baseInfo.WorkSheets;
            for (int i = 0; i < this.baseInfo.WorkSheets.Count; ++i)
            {
                dynamic worksheet = worksheets[i + 1];
                list.Add(new WorkSheetModel(worksheet));
            }
        }

        public Newtonsoft.Json.Linq.JObject ToJsonObject()
        {
            return new JObject(
                new JProperty("ExcelJsonFormatInfo",
                    new JArray(
                        from o in list
                        select o.ToJsonObject()
                        )
                    )
                );
        }
    }
}
