using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JssonPortAddIn.Model.ExcelModel
{
    interface IConvertableJson
    {
        JToken ToJsonObject();
    }
}
