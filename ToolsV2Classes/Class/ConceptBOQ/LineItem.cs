using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ToolsV2Classes.Class.ConceptBOQ
{
    public class LineItem
    {
        public string IdentityInfo { get; set; }//From Excel
        public int ItemRow { get; set; }//row number
        public string RevitParameter { get; set; }//Parameter name (type mark/type name/family name)
        public string ParameterValue { get; set; }//ID of aboce parameter
        public string CategoryValue { get; set; }//Category of line item
        public string ItemUnit { get; set; }//unit of line item
        public string ItemQty { get; set; }//quantity of line item from Revit

        public LineItem CreateLineItem(string idInfo, int row, string unit, string cat)
        {
            this.IdentityInfo = idInfo;
            this.ItemRow = row;
            this.ItemUnit = unit;
            this.CategoryValue = cat;

            GetRevitParameterName(this, idInfo);

            return this;
        }

        public void GetRevitParameterName(LineItem li, string idInfo)
        {
            int totalChar = idInfo.Length;


            int index_ = idInfo.IndexOf("_");
            if (index_ > 0)
            {
                li.RevitParameter = idInfo.Substring(0, index_);
                li.ParameterValue = idInfo.Substring(index_ + 1, totalChar - 1 - index_);
            }

        }


    }
}
