using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using DeskAutomation.Helper_Classes;

namespace DeskAutomation.Helper_Classes
{
    public class DoorData
    {
        public FamilyInstance Door { get; set; }

        public Room EnterRoom { get; set; }
        public Room ExitRoom { get; set; }



        public void GetDoorInfo(FamilyInstance door)
        {
            Door = door;
            EnterRoom = door.ToRoom;
            ExitRoom = door.FromRoom;
        }


    }

}
