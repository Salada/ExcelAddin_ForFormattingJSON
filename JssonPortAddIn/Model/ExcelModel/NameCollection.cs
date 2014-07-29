using System.Collections.Generic;

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

        public IEnumerable<string> Select()
        {
            for (int i = 0; i < rowNames.Length; ++i)
            {
                yield return this[i];
            }
        }

        public int Length { get { return rowNames.Length; }  }
    }
}
