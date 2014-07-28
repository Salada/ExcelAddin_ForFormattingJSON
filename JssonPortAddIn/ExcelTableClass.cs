using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace JssonPortAddIn
{
    class ExcelTableClass
    {
        private dynamic worksheet; // this is com object. then, must only use dynamic type.
        
        private List<SimpleExpandoObject> dataObjects;
        public string CodeName
        {
            get;
            set;
        }
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

        public string SheetName
        {
            get;
            set;
        }

        public ExcelTableClass()
        {
            throw new NotImplementedException();
        }


        public ExcelTableClass(dynamic worksheet)
        {
            this.worksheet = worksheet;
            this.dataObjects = new List<SimpleExpandoObject>();
            
            initializeObject();
        }

        private void initializeObject()
        {
            this.SheetName = this.worksheet.Name;
            this.CodeName = this.worksheet.CodeName;
            this.TableName = this.worksheet.Cells.ListObject.Name;


            dynamic tableRange = this.worksheet.Cells.ListObject.Range;
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

        public void ToJsonObject()
        {
            
        }
    }
}
