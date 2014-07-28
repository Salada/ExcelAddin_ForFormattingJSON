using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JssonPortAddIn
{
    class NameCollection
    {
        private dynamic rowNames;

        public string this[int index]
        {
            get
            {
                return rowNames[1, index];
            }
            set
            {
                rowNames[1, index] = value;
            }
        }

        public NameCollection(dynamic rowNames)
        {
            this.rowNames = rowNames.Clone();
        }
    }
}
