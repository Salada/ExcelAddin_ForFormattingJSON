using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JssonPortAddIn
{
    class SimpleExpandoObject : DynamicObject
    {
        private Dictionary<object, object> vals = new Dictionary<object, object>();

        public override bool TryGetMember(System.Dynamic.GetMemberBinder binder, out object result)
        {
            return vals.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(System.Dynamic.SetMemberBinder binder, object value)
        {
            vals[binder.Name] = value;
            return true;
        }

        public bool Add(string key, object value)
        {
            if (!vals.ContainsKey(key))
            {
                vals.Add(key, value);
                return true;
            }
            else
            {
                return false;
            }
        }

        internal JObject ToJsonObject()
        {
            return new JObject(
                from v in vals
                select new JProperty((string)v.Key, v.Value)
                );
        }
    }
}
