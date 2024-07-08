using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOQStandardCheck
{
    public class SheetInfoClass
    {

        public string sheetName { get; set; }
        public List<rowInfoClass> rowInfoClassList { get; set; }
        // public List<string> description { get; set; }
    }

    public class rowInfoClass
    {
        public string rowNumber { get; set; }
        public string labelValue { get; set; }

        public string descriptionValue { get; set; }
        public string omniCodeValue { get; set; }

        public bool isMerged { get; set; }


    }
}
