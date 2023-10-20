using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class PowerAndData : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            try
            {
                #region Get Rooms from Selection
                //Collect the rooms from user selection, filter out unrequired elements
                List<Room> roomList = new List<Room>();
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (0 == selectedIds.Count)
                {
                    // If no elements are selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + " You haven't selected any Rooms!");
                    goto skipTool;
                }
                else
                {
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = uidoc.Document.GetElement(id);
                        if (elem is Room)
                        {
                            Room testRedundant = elem as Room;
                            string programType = LightingToolMethods.GetParamVal(testRedundant, "WW-ProgramType");
                            if (testRedundant.Area != 0 && programType.ToLower() == "work")
                            {
                                //Add rooms to roomList
                                roomList.Add(elem as Room);
                            }
                        }
                    }
                    if (0 == roomList.Count)
                    {
                        // If no rooms are selected.
                        TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Rooms!");
                        goto skipTool;
                    }
                }
                #endregion

                #region Ask user to delete existing grommets
                ToolsV2Classes.Class.Power_and_Data.DeleteExisting form2 = new ToolsV2Classes.Class.Power_and_Data.DeleteExisting(commandData);
                form2.ShowDialog();
                bool wantToDelete = form2.grommetDelete;
                #endregion

                foreach (Room r1 in roomList)
                {
                    //Delete Existing Grommets in the selected rooms
                    if (!wantToDelete)
                    {
                        using (Transaction trans1 = new Transaction(doc, "Delete Existing Grommets"))
                        {
                            trans1.Start();
                            DeleteExistingGrommets(r1);
                            trans1.Commit();
                        }
                    }

                    #region Finding Floor Below Room Center
                    //Find floor below the room
                    Reference floorReferenceBelow = FindFloorBelow(doc, r1);
                    if(floorReferenceBelow == null)
                    {
                        //If floor not found, use level as default placement reference plane
                        floorReferenceBelow = r1.Level.GetPlaneReference();
                    }
                    #endregion

                    //Unordered list of desks in deskObjects
                    List<Desk> deskObjects = new List<Desk>();
                    //Retrieving desks from selected room
                    List<FamilyInstance> deskList = GetDesks(r1);
                    foreach (FamilyInstance d in deskList)
                    {
                        //Creating a database for each desk with required properties in Desk class
                        Desk d1 = new Desk();
                        d1.DeskParameters(d);
                        deskObjects.Add(d1);
                        /*
                        Getting transformed coordinates for each desk based on orientation.
                        The transformation will give an virtual desk layout where all the desks are in either 0/180/360 degree orientation
                        */
                        TransformCoordinates(d1);
                    }

                    //globalDesks. Each list of this variable will contain sorted desks
                    List<List<Desk>> globalDesks = new List<List<Desk>>();

                    //Sorting and Grouping desks as per orientation or angle of desk placement
                    var deskOriented = deskObjects.OrderBy(p => p.AnglePair).GroupBy(p => p.AnglePair).Select(grp => grp.ToList()).ToList();

                    foreach (var deskOrientedGrp in deskOriented)
                    {
                        //Sorting and Grouping Desks as per transformed X Coordinate
                        var deskXOriented = deskOrientedGrp.OrderBy(p => p.TransX).GroupBy(p => p.TransX).Select(grp => grp.ToList()).ToList();
                        foreach (var deskXGrp in deskXOriented)
                        {
                            //Splitting grouped desks if the distance between transformed Y coordinate is more that 5 feet
                            globalDesks.AddRange(GroupSplitter(deskXGrp));
                        }
                    }

                    List<DeskColumn> deskColumns = new List<DeskColumn>();
                    foreach (var col in globalDesks)
                    {
                        //Creating a database for each desk group/column with required properties in DeskColumn class
                        DeskColumn dc1 = new DeskColumn();
                        dc1.DeskColumnPara(col);
                        deskColumns.Add(dc1);
                    }

                    foreach (DeskColumn dc1 in deskColumns)
                    {
                        if(dc1.GroupType == "single")
                        {
                            //Get power and data grommet type, numbers and coordinate for each single desk group/column 
                            GetLayoutSingleColumn(dc1);
                        }
                        if (dc1.GroupType == "double")
                        {
                            //Get power and data grommet type, numbers and coordinate for each double desks group/column 
                            GetLayoutDoubleColumn(dc1);
                        }
                    }

                    //Get the Family types of the grommets
                    List<FamilySymbol> grommetTypeList = new List<FamilySymbol>
                    {
                        GetFamilyType("p1d1", doc),
                        GetFamilyType("p2d2", doc),
                        GetFamilyType("p3d3", doc),
                        GetFamilyType("p4d4", doc)
                    };

                    using (Transaction trans2 = new Transaction(doc, "Power and Data"))
                    {
                        trans2.Start();
                        foreach (DeskColumn dc1 in deskColumns)
                        {
                            foreach (FamilySymbol type in grommetTypeList)
                            {
                                //Activate the family type if not activated already
                                if(!type.IsActive)
                                {
                                    type.Activate();
                                }
                            }
                            //Place grommets
                            PlaceGrommets(dc1, doc, grommetTypeList, floorReferenceBelow);
                        }
                        trans2.Commit();
                    }
                }
            skipTool:
                // Return success result

                string toolName = "Grommet";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Grommet";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }
        }

        static void DeleteExistingGrommets (Room room)
        {
            //Delete existing grommets within the selected room
            Document doc = room.Document;
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
            FilteredElementCollector familyInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_DataDevices)
                .WherePasses(filter);
            int roomid = room.Id.IntegerValue;
            List<ElementId> existingGrommets = new List<ElementId>();
            foreach (FamilyInstance fi in familyInstances)
            {
                if (null != fi.Room
                    && fi.Room.Id.IntegerValue.Equals(roomid)
                    && (fi.Symbol.FamilyName.ToLower().Contains("powerdata")
                    || fi.Symbol.FamilyName.ToLower().Contains("poweranddata")))
                {
                    existingGrommets.Add(fi.Id);
                }
                else if (fi.Symbol.FamilyName.ToLower().Contains("powerdata_floorbased"))
                {
                    existingGrommets.Add(fi.Id);
                }
            }
            if (existingGrommets.Count() > 0)
            {
                foreach (ElementId fiId in existingGrommets)
                {
                    doc.Delete(fiId);
                }
            }
        }
        static Reference FindFloorBelow(Document doc, Room room)
        {
            /*
            Find the floor below the Room bounding box center point.
            This may fail sometimes if the room is C shaped or even longer L shaped.
            If the room bounding box center falls outside the room, we'll use Level as default reference.
            */
            Reference reference = null;

            BoundingBoxXYZ roomBox = room.get_BoundingBox(null);
            XYZ center = roomBox.Min.Add(roomBox.Max).Multiply(0.5);
            if (room.IsPointInRoom(center))
            {
                int roomLevelId = room.Level.Id.IntegerValue;
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                bool isNotTemplate(View3D v3) => !(v3.IsTemplate);
                View3D view3D = collector.OfClass(typeof(View3D)).Cast<View3D>().First<View3D>(isNotTemplate);
                XYZ rayDirection = new XYZ(0, 0, -1);
                ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                try
                {
                    ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext = refIntersector.FindNearest(center, rayDirection);
                    reference = referenceWithContext.GetReference();
                    var intersection = reference.ElementId;
                    var floor = doc.GetElement(intersection);
                    int floorLevelId = floor.LevelId.IntegerValue;
                    if (roomLevelId != floorLevelId)
                    {
                        reference = null;
                        int error = roomLevelId / 0; //Deliberately creating exception
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Revit", "Floor below Room not found, Level will be considered as Grommet Host for placement.");
                }
            }
            else
            {
                TaskDialog.Show("Revit", "Geometric Center for one or multiple room may lie outside the room volume."
                    + Environment.NewLine + "Level will be considered as Grommet Host for placement.");
                reference = null;
            }
            return reference;
        }
        static FamilySymbol GetFamilyType (string typeName, Document doc)
        {
            //Returns the family type of the Grommet family based on input
            FamilySymbol grommetType = null;
            FilteredElementCollector dataDevices = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DataDevices).WhereElementIsElementType();

            grommetType = dataDevices.OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.ToLower().Contains("wwi-poweranddata-grommet"))
                .Where(x => x.Name.ToLower().Contains(typeName)).FirstOrDefault();

            //If grommet family not found in project, Load the family from google drive
            if (grommetType == null)
            {
                TaskDialog.Show("Revit", "Standard Grommet family is not loaded into the project."
                    + Environment.NewLine
                    + "Tool will try to load the latest Grommet fixture Family : 'WWI-PowerAndData-Grommet'");
                using (Transaction tx = new Transaction(doc, "Load Grommet Family"))
                {
                    tx.Start();
                    if (grommetType == null)
                    {
                        string path = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Data Devices\\WWI-PowerAndData-Grommet.rfa";
                        FamilyLoadOption newOption = new FamilyLoadOption();
                        doc.LoadFamily(path, newOption, out Family grommetFamily);
                        if (grommetFamily == null)
                        {
                            TaskDialog.Show("Revit", "Grommet Family cannot be Loaded, please load the family 'WWI-PowerAndData-Grommet' manually :(");
                        }
                        else
                        {
                            TaskDialog.Show("Revit", "Grommet Family Loaded: " + grommetFamily.Name + Environment.NewLine + " :) ");
                        }
                    }
                    tx.Commit();
                    grommetType = dataDevices.OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                        .Where(x => x.FamilyName.ToLower().Contains("wwi-poweranddata-grommet"))
                        .Where(x => x.Name.ToLower().Contains(typeName)).FirstOrDefault();
                }
            }
            return grommetType;
        }
        static List<FamilyInstance> GetDesks(Room room)
        {
            /*             
            Returns list of 1_person-office-desk family instances from selected room.
            Ghosted Desks will be ignored. Hot desks are not considered yet.
            */
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter filter
              = new BoundingBoxIntersectsFilter(outline);
            Document doc = room.Document;
            FilteredElementCollector familyInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_FurnitureSystems)
                .WherePasses(filter);
            int roomid = room.Id.IntegerValue;
            List<FamilyInstance> a = new List<FamilyInstance>();
            foreach (FamilyInstance fi in familyInstances)
            {
                if (null != fi.Room 
                    && fi.Room.Id.IntegerValue.Equals(roomid) 
                    && fi.Symbol.FamilyName.ToLower().Contains("1_person-")
                    && fi.get_Parameter(new Guid("afbfd170-9396-4faf-bd9d-6d03aae40976")).AsValueString().Equals("No"))
                {
                    a.Add(fi);
                }
            }
            return a;
        }
        static void TransformCoordinates(Desk d)
        {
            //Transforming the coordinates of desks to get all desks as set vertical columns
            if (d.AnglePair == 0)
            {
                d.TransX = d.RoundedX;
                d.TransY = d.RoundedY;
            }
            else
            {
                var t_form = Transform.CreateRotation(XYZ.BasisZ, d.AnglePair * (Math.PI / 180));
                d.TransX = Math.Round(t_form.OfPoint(d.Location.Point).X);
                d.TransY = Math.Round(t_form.OfPoint(d.Location.Point).Y);
            }
        }
        static XYZ TransformGrommetsSingle(Desk d1)
        {
            //Transform grommet location from Desk location point.
            XYZ deskPoint = d1.XYZPoint;
            XYZ facingDir = d1.FacingPoint;
            facingDir = facingDir.Negate();
            Transform t_form_1 = Transform.CreateTranslation(facingDir);
            XYZ grommetMidPoint = t_form_1.OfPoint(deskPoint);
            XYZ zPoint = new XYZ(grommetMidPoint.X, grommetMidPoint.Y, grommetMidPoint.Z + 1);
            XYZ zAxis = (zPoint - grommetMidPoint).Normalize();
            Transform t_form_2;
            if (d1.AngleDegree == 180 || d1.AngleDegree == 270)
            {
                t_form_2 = Transform.CreateRotationAtPoint(zAxis, Math.PI * 0.5, grommetMidPoint);
            }
            else
            {
                t_form_2 = Transform.CreateRotationAtPoint(zAxis, Math.PI * 1.5, grommetMidPoint);
            }
            XYZ grommetPoint = t_form_2.OfPoint(deskPoint);
            return grommetPoint;
        }
        static XYZ TransformGrommetsDouble(Desk d1)
        {
            //Transform grommet location from Desk location point.
            XYZ deskPoint = d1.XYZPoint;
            double angle = d1.AngleRad;
            XYZ grommetPoint = new XYZ(deskPoint.X - Math.Abs(Math.Sin(angle)), deskPoint.Y - Math.Abs(Math.Cos(angle)), deskPoint.Z);

            return grommetPoint;
        }
        static void GetLayoutSingleColumn (DeskColumn dc1)
        {
            //Generate the layout for grommets for single row/column of desk
            //Type of grommet and Location is stored in DeskColumn object property
            int deskCount = dc1.DeskCol.Count();
            int leftover = deskCount % 3;
            int groupsOf3 = deskCount / 3;
            int counter = 0;
            int tempCounter = 0;
            List<XYZ> grommetsXYZ = new List<XYZ>();
            List<string> PnD = new List<string>();
            char orientation = 'v';
            if(dc1.DeskCol[0].AnglePair == 90)
            {
                orientation = 'h';
            }
            if (leftover == 1 && groupsOf3 > 0)
            {
                PnD.Add("P2D2");
                grommetsXYZ.Add(TransformGrommetsSingle(dc1.DeskCol[counter+1]));
                counter = 2;
            }

            for (int i = counter; i < deskCount; i++)
            {
                tempCounter++;
                if (tempCounter == 3)
                {
                    PnD.Add("P3D3");
                    grommetsXYZ.Add(TransformGrommetsSingle(dc1.DeskCol[i-1]));
                    tempCounter = 0;
                }
            }
            if (tempCounter != 0)
            {
                PnD.Add(GetPDNumbers(tempCounter));
                grommetsXYZ.Add(TransformGrommetsSingle(dc1.DeskCol[deskCount-1]));
            }
            dc1.PnD = PnD;
            dc1.GrommetsXYZ = grommetsXYZ;
            dc1.GrommetOrientation = orientation;
        }
        static void GetLayoutDoubleColumn (DeskColumn dc1)
        {
            //Generate the layout for grommets for double row/column of desk
            //Type of grommet and Location is stored in DeskColumn object property
            var yGroupings = dc1.DeskCol.OrderBy(p => p.TransY).GroupBy(p => p.TransY).Select(grp => grp.ToList()).ToList();
            List<string> PnD = new List<string>();
            List<XYZ> grommetsXYZ = new List<XYZ>();
            int tempDeskCount = 0;
            int tempGroupCount = 0;
            var lastGroup = yGroupings[0];
            char orientation = 'h';
            if (dc1.DeskCol[0].AnglePair == 90)
            {
                orientation = 'v';
            }
            foreach (var grp1 in yGroupings)
            {
                tempDeskCount += grp1.Count();
                tempGroupCount++;
                if(tempGroupCount == 2)
                {
                    PnD.Add(GetPDNumbers(tempDeskCount));
                    grommetsXYZ.Add(TransformGrommetsDouble(grp1[0]));
                    tempGroupCount = 0;
                    tempDeskCount = 0;
                }
                lastGroup = grp1;
            }
            if (tempDeskCount != 0)
            {
                PnD.Add(GetPDNumbers(tempDeskCount));
                grommetsXYZ.Add(TransformGrommetsDouble(lastGroup[0]));
            }
            dc1.PnD = PnD;
            dc1.GrommetsXYZ = grommetsXYZ;
            dc1.GrommetOrientation = orientation;
        }
        static string GetPDNumbers (int deskCount)
        {
            //Intermadiate function for grommet type selection based on number of desks per group.
            string PnD = "";
            switch (deskCount)
            {
                case 0:
                    break;
                case 1:
                    PnD = "P1D1";
                    break;
                case 2:
                    PnD = "P2D2";
                    break;
                case 3:
                    PnD = "P3D3";
                    break;
                case 4:
                    PnD = "P4D4";
                    break;
                default:
                    TaskDialog.Show("Revit", "Default case statement error");
                    break;
            }
            return PnD;
        }
        static List<List<Desk>> GroupSplitter (List<Desk> inputList)
        {
            /*
             * Splits the desks of a single Column if they are more than 5 feet apart.
             * This will return a List of List of Desks and append to globalDesks List of List of desks
             * If there is no split in Y direction, it'll just return List of List of Desks containing only one List of Desk. Confusing enough??
            */
            List<List<Desk>> yGrouped = new List<List<Desk>>();
            List<Desk> tempList = new List<Desk>();
            inputList = inputList.OrderBy(p => p.TransY).ToList();

            double d1Y = inputList.FirstOrDefault().TransY;
            double d2Y = 0;
            foreach ( Desk d1 in inputList)
            {
                d2Y = d1.TransY;
                double diff = Math.Abs(d1Y - d2Y);
                if (diff < 5)
                {
                    tempList.Add(d1);
                }
                else
                {
                    yGrouped.Add(tempList);
                    tempList = new List<Desk>{d1};
                }
                d1Y = d2Y;
            }
            if (tempList.Count > 0)
            {
                yGrouped.Add(tempList);
            }
            return yGrouped;
        }
        static void PlaceGrommets(DeskColumn dc1, Document doc, List<FamilySymbol> grommetTypeList, Reference floorReferenceBelow)
        {
            //Places grommets based based on DeskColumn properties -- Location & Type of Grommet
            //If floor reference is found below room, the host will be floor, otherwise default host is room level
            List<XYZ> grommetsXYZ = dc1.GrommetsXYZ;
            List<string> grommetsType = dc1.PnD;
            int indexOf = 0;
            XYZ refDir = new XYZ(1, 0, 0);
            FamilyInstance grommetPlaced = null;
            foreach (string grommet in grommetsType)
            {
                switch (grommet)
                {
                    case "P1D1":
                        grommetPlaced = doc.Create.NewFamilyInstance(floorReferenceBelow, grommetsXYZ[indexOf], refDir, grommetTypeList[0]);
                        break;
                    case "P2D2":
                        grommetPlaced = doc.Create.NewFamilyInstance(floorReferenceBelow, grommetsXYZ[indexOf], refDir, grommetTypeList[1]);
                        break;
                    case "P3D3":
                        grommetPlaced = doc.Create.NewFamilyInstance(floorReferenceBelow, grommetsXYZ[indexOf], refDir, grommetTypeList[2]);
                        break;
                    case "P4D4":
                        grommetPlaced = doc.Create.NewFamilyInstance(floorReferenceBelow, grommetsXYZ[indexOf], refDir, grommetTypeList[3]);
                        break;
                    default:
                        TaskDialog.Show("Revit", "Default Switch Statement Error");
                        break;
                }
                indexOf++;
                if (grommetPlaced != null)
                {
                    ChangeGrommetAnnotation(grommetPlaced, dc1.GrommetOrientation);
                }
            }
        }
        static void ChangeGrommetAnnotation (FamilyInstance grommet, char orientation)
        {
            Parameter hSymbol = grommet.GetParameters("Horizontal Symbol").FirstOrDefault();
            Parameter vSymbol = grommet.GetParameters("Vertical Symbol").FirstOrDefault();
            Parameter symbolOffset = grommet.GetParameters("Symbol Offset").FirstOrDefault();
            Parameter grommetOffset = grommet.GetParameters("Offset").FirstOrDefault();
            symbolOffset.Set(1);
            grommetOffset.Set(0);
            if (orientation == 'v')
            {
                hSymbol.Set(0);
                vSymbol.Set(1);
            }
            else
            {
                hSymbol.Set(1);
                vSymbol.Set(0);
            }
        }

    }

    public class DeskColumn
    {
        //Class for creating desk column objects and it's properties

        public List<Desk> DeskCol { get; set; } // List containing Desk objects
        public string GroupType { get; set; } // Either Single or Double layout of desks
        public List<string> PnD { get; set; } // List of type Grommet to be placed for the desk column/row
        public List<XYZ> GrommetsXYZ { get; set; } // List of XYZ points for Grommet placement
        public char GrommetOrientation { get; set; } //To change the grommet annotation orientation
        public void DeskColumnPara (List<Desk> inputList)
        {
            DeskCol = inputList;
            var yGroupings = DeskCol.OrderBy(p => p.TransY).GroupBy(p => p.TransY).Select(grp => grp.ToList());
            double remin = (double)DeskCol.Count() / (double)yGroupings.Count();
            if (remin <= 1)
            {
                GroupType = "single";
            }
            else
            {
                GroupType = "double";
            }
        }
    }

    public class Desk
    {
        // Class for creating desk as objects and create properties related to Desk family instance

        public void DeskParameters (FamilyInstance desk)
        {
            DeskElem = desk;
            Location = desk.Location as LocationPoint;
            XYZPoint = Location.Point as XYZ;
            RoundedX = Math.Round(Location.Point.X);
            RoundedY = Math.Round(Location.Point.Y);
            FacingPoint = desk.FacingOrientation;
            AngleRad = desk.FacingOrientation.AngleOnPlaneTo(XYZ.BasisX, XYZ.BasisZ);
            AngleDegree = Math.Round(desk.FacingOrientation.AngleOnPlaneTo(XYZ.BasisX, XYZ.BasisZ)*180/Math.PI);
            if ( AngleDegree == 0 || AngleDegree == 180 || AngleDegree == 360 )
            {
                AnglePair = 0;
            }
            else if (AngleDegree < 180)
            {
                AnglePair = AngleDegree;
            }
            else if (AngleDegree > 180)
            {
                AnglePair = AngleDegree - 180;
            }
        }

        public FamilyInstance DeskElem { get; set; } //Desk as a family instance
        public LocationPoint Location { get; set; } //Location of the desk
        public XYZ XYZPoint { get; set; } //XYZ point of Desk
        public XYZ FacingPoint { get; set; } //Facing direction of Desk
        public double RoundedX { get; set; } //X coordinate of desk rounded to 1 foot
        public double RoundedY { get; set; } //Y coordinate of desk rounded to 1 foot
        public double TransX { get; set; } //Transformed X coordinate of desk rounded to 1 foot
        public double TransY { get; set; } //Transformed Y coordinate of desk rounded to 1 foot
        public double AngleRad { get; set; } //Facing angle of Desk in Radians
        public double AngleDegree { get; set; } //Facing angle of Desk in Degrees rounded to 1 degree
        public double AnglePair { get; set; } //Pairing property of the desks with opposite angle to each other

    }
}