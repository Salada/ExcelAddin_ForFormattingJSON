using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JssonPortAddIn.Model.ExcelModel;

namespace JssonPortAddIn
{
    class ExcelTableModel : IConvertableJson
    {
        private WorkSheetModel parentModel; // this is com object. then, must only use dynamic type.
        
        private List<SimpleExpandoObject> dataObjects;

        public NameCollection RowNames
        {
            get;
            private set;
        }

        public string TableName
        {
            get;
            set;
        }


        public ExcelTableModel()
        {
            this.dataObjects = new List<SimpleExpandoObject>();
        }


        public ExcelTableModel(WorkSheetModel parentModel) : this()
        {
            this.parentModel = parentModel;

            initializeObject();
        }

        private void initializeObject()
        {
            this.TableName = this.parentModel.WorkSheet.Cells.ListObject.Name;

            dynamic tableRange = this.parentModel.WorkSheet.Cells.ListObject.Range;
            decimal count = 1;

            foreach (dynamic row_data in tableRange.Rows)
            {
                dynamic rows = row_data.FormulaLocal;
                if (count == 1)
                {
                    this.RowNames = new NameCollection(rows);
                }
                else
                {
                    SimpleExpandoObject dynObject = new SimpleExpandoObject();
                    for (int i = 1; i <= rows.Length; ++i)
                    {
                        dynObject.Add(this.RowNames[i], rows[1, i]);
                    }
                    dataObjects.Add(dynObject);
                }

                count++;
            }
        }

        public JObject ToJsonObject()
        {
            JObject rss = new JObject();

            JArray dataArr = new JArray(from p in this.dataObjects 
                                        select p.ToJsonObject());

            rss.Add(new JProperty(TableName, dataArr));
            return rss;
        }
            
    }
}
