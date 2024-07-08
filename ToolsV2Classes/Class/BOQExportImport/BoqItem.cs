using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ToolsV2Classes
{
    public class BoqItem
    {
        public BoqItem(Element typeName, string typeMark)
        {
            this.typeName = typeName;
            this.typeMark = typeMark;
            this.costAvailable = false;
            this.cost = 0;
        }

        public Element typeName { get; set; }
        public string typeMark { get; set; }
        public int cost { get; set; }
        public bool costAvailable { get; set; }
    }
}
