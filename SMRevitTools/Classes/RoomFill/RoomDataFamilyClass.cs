using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMRevitTools.Classes.RoomFill
{
    internal class RoomDataFamilyClass
    {
        public FamilyInstance RoomDataFamilyInstace { get; set; }

        public string RoomDataFamilyName { get; set; }

        public string RoomNameSet { get; set; }

        public XYZ CenterLocation { get; set; }

        public UV CenterLocationUV { get; set; }

        public Level RoomDataLevel { get; set; }

        public void GetFamilyInstanceData (FamilyInstance fi, Document doc)
        {
            this.RoomDataFamilyInstace = fi;
            this.RoomDataFamilyName = fi.Symbol.FamilyName;
            this.RoomNameSet = RoomDataFamilyName.Substring(12, RoomDataFamilyName.Length - 1-11);
            this.CenterLocation = (fi.Location as LocationPoint).Point;
            this.RoomDataLevel = doc.GetElement(fi.LevelId) as Level;
            this.CenterLocationUV = new UV(CenterLocation.X, CenterLocation.Y);

        }

        public void GetWKInstanceData(FamilyInstance fi, Document doc)
        {
            this.RoomDataFamilyInstace = fi;
            this.RoomDataFamilyName = fi.Symbol.FamilyName;
            //this.RoomNameSet = RoomDataFamilyName.Substring(12, RoomDataFamilyName.Length - 1 - 11);
            this.CenterLocation = (fi.Location as LocationPoint).Point;
            this.RoomDataLevel = doc.GetElement(fi.LevelId) as Level;
            this.CenterLocationUV = new UV(CenterLocation.X, CenterLocation.Y);

        }

    }
}
