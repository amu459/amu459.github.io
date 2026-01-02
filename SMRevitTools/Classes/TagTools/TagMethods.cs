using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq.Expressions;

namespace SMRevitTools.Classes.TagTools
{
    public class TagMethods
    {
        public static List<FamilyInstance> GetDoorWithACDFRD(List<FamilyInstance> doors) 
        { 
            List<FamilyInstance> result = new List<FamilyInstance>();

            foreach (FamilyInstance door in doors)
            {
                //get parameter ACD
                string acdParam = door.LookupParameter("SM-Door ACD").AsString();
                //get parameter FRD
                string frdParam = door.LookupParameter("SM-Door FRD").AsString();

                //if either of them are not null add to result
                if(acdParam != null ||  frdParam != null)
                {
                    result.Add(door);
                }
            }
            return result;
        }






    }


}
