using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOQStandardCheck
{
    public class CompareReportClass
    {

        public List<string> NotFoundList { get; set; }
        public List<string> NewlyAddedList { get; set; }
        public List<string> NotComparedList { get; set; }
        public string sheetName { get; set; }
    }

    public class ErrorData
    {
        public List<CompareReportClass> ComparedReportList { get; set; }

        public List<string> otherErrorList { get; set; }

    }

}
