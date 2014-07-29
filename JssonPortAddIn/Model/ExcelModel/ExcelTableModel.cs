using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JssonPortAddIn.Model.ExcelModel;
using Excel = Microsoft.Office.Tools.Excel;

namespace JssonPortAddIn
{
    class ExcelTableModel : IConvertableJson
    {
        private ExcelWorkSheetModel parentModel; // this is com object. then, must only use dynamic type.
        
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


        public ExcelTableModel(ExcelWorkSheetModel parentModel) : this()
        {
            this.parentModel = parentModel;

            initializeObject();
        }

        public ExcelTableModel(ExcelWorkSheetModel parentModel, JObject obj) : this()
        {
            
            this.parentModel = parentModel;
            this.TableName = obj["TableName"].ToString();

            decimal row = 1;
            decimal col = 1;

            decimal rowBase = row;
            decimal colBase = col;

            decimal colLast = col;
            decimal rowLast = row;

            
            dynamic worksheet = this.parentModel.WorkSheet;
            dynamic application = worksheet.Application;

            foreach (JObject iterableObj in obj["Items"])
            {
                if (row == rowBase)
                {
                    foreach (JProperty property in iterableObj.Properties())
                    {
                        worksheet.Cells[row, col].FormulaLocal = property.Name;
                        col++;
                    }
                    colLast = col - 1;
                    row++; col = colBase;
                }

                foreach (JProperty property in iterableObj.Properties())
                {
                    worksheet.Cells[row, col].FormulaLocal = property.Value.ToString();
                    col++;
                }
                row++; col = colBase;
            }
            rowLast = row - 1;

            dynamic activeSheet = worksheet.Application.ActiveSheet;
            worksheet.ListObjects.Add(
                Source : activeSheet.Range(activeSheet.Cells(rowBase, colBase), activeSheet.Cells(rowLast, colLast)) // Range
                ).Name = this.TableName;
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

        public JToken ToJsonObject()
        {
            return new JObject(
                new JProperty("TableName", TableName),
                new JProperty("Items", new JArray(from p in this.dataObjects 
                                                    select p.ToJsonObject()
                                )
                            )
                );
        }
            
    }
}
