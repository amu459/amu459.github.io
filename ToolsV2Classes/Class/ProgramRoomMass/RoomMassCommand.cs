#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using System.Diagnostics;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI.Selection;
#endregion // Namespaces

namespace ToolsV2Classes
{

    // 3D Rooms Création
    [Transaction(TransactionMode.Manual)]
    public class RoomMassCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view;
            view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            int roomCount = 0;
            List<ElementId> failedRooms = new List<ElementId>();

            // Deleting existing DirectShape
            // get ready to filter across just the elements visible in a view 
            FilteredElementCollector coll = new FilteredElementCollector(doc, view.Id);
            coll.OfClass(typeof(DirectShape));
            IEnumerable<DirectShape> DSdelete = coll.Cast<DirectShape>();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Delete elements");
                try
                {
                    foreach (DirectShape ds in DSdelete)
                    {
                        ICollection<ElementId> ids = doc.Delete(ds.Id);
                    }
                    tx.Commit();
                }
                catch (ArgumentException)
                {
                    tx.RollBack();
                }
            }

            //Collect all the rooms
            #region Get All Rooms except Container level and unplaced rooms

            List<Room> roomList = new List<Room>();
            var roomIList = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(SpatialElement))
                .Where(e => e.GetType() == typeof(Room))
                .Cast<Room>();

            foreach(Room r in roomIList)
            {
                if(r.Level.Elevation < -10 || r.Area < 1)
                {
                    continue;
                }
                else
                {
                    roomList.Add(r);
                }
            }

            if (0 == roomList.Count)
            {
                // If no rooms found.
                TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine 
                    + " I couldn't find any Rooms!" + Environment.NewLine
                    + " Let's go to Blr thindis and chill!");
            }
            #endregion


            using (Transaction tr = new Transaction(doc))
            {
                tr.Start("Create Mass");
                //  Iterate the list and gather a list of boundaries
                foreach (Room room in roomList)
                {
                    String _family_name = "testRoom-" + room.UniqueId.ToString();

                    // Found BBOX
                    BoundingBoxXYZ bb = room.get_BoundingBox(null);
                    XYZ pt = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, bb.Min.Z);
                    SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions
                    {
                        SpatialElementBoundaryLocation =
                      SpatialElementBoundaryLocation.Center
                    };
                    //  Get the room boundary
                    IList<IList<BoundarySegment>> boundaries =
                        room.GetBoundarySegments(opt); // 2012

                    // a room may have a null boundary property:
                    int n = 0;
                    if (null != boundaries)
                    {
                        n = boundaries.Count;
                    }

                    //  Iterate to gather the curve objects
                    TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
                    builder.OpenConnectedFaceSet(true);

                    // Add Direct Shape
                    List<CurveLoop> curveLoopList = new List<CurveLoop>();
                    if (0 < n)
                    {
                        foreach (IList<BoundarySegment> b in boundaries)
                        {
                            List<Curve> profile = new List<Curve>();
                            foreach (BoundarySegment s in b)
                            {
                                Curve curve = s.GetCurve();
                                profile.Add(curve); //add shape for instant object
                            }
                            try
                            {
                                CurveLoop curveLoop = CurveLoop.Create(profile);
                                curveLoopList.Add(curveLoop);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                                failedRooms.Add(room.Id);
                            }
                        }
                    }

                    try
                    {
                        SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);

                        Frame frame = new Frame(pt, XYZ.BasisX, -XYZ.BasisZ, XYZ.BasisY);

                        //  Simple insertion point
                        XYZ pt1 = new XYZ(0, 0, 0);
                        //  Our normal point that points the extrusion directly up
                        XYZ ptNormal = new XYZ(0, 0, 2400);
                        //  The plane to extrude the mass from
                        Plane m_Plane = Plane.CreateByNormalAndOrigin(ptNormal, pt1);
                        // SketchPlane m_SketchPlane = m_FamDoc.FamilyCreate.NewSketchPlane(m_Plane);
                        SketchPlane m_SketchPlane = SketchPlane.Create(doc, m_Plane); // 2014

                        Solid roomSolid;

                        roomSolid = GeometryCreationUtilities
                            .CreateExtrusionGeometry(curveLoopList, ptNormal, 8);

                        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_Mass));

                        ds.SetShape(new GeometryObject[] { roomSolid });
                        ds.SetName(_family_name);
                        string programType = LightingToolMethods
                            .GetParamVal(room, "WW-ProgramType");
                        ColorScheme colorSch = new ColorScheme(programType);
                        Color color = colorSch.actualColor;
                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        ogs.SetProjectionLineColor(color);
                        ogs.SetSurfaceForegroundPatternColor(color);
                        ElementId solidId = new ElementId(19);
                        ogs.SetSurfaceForegroundPatternId(solidId);
                        ogs.SetCutLineColor(color);
                        ogs.SetCutForegroundPatternColor(color);
                        ogs.SetCutForegroundPatternId(solidId);

                        doc.ActiveView.SetElementOverrides(ds.Id, ogs);
                        roomCount++;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        failedRooms.Add(room.Id);
                    }
                }
                tr.Commit();
            }
            string failedMessage = "Total Mass Created = " + roomCount.ToString() + 
               Environment.NewLine + "No of failed Rooms = " + failedRooms.Count() +
               Environment.NewLine + "Failed Rooms' Element Ids:";
            if (failedRooms.Count() > 0)
            {
                foreach (ElementId rId in failedRooms)
                {
                    failedMessage += Environment.NewLine + rId.ToString();
                }
            }
            else
            {
                failedMessage += " None :)";
            }
            TaskDialog.Show("3D rOOOOms:", failedMessage);

            string toolName = "3DrOOms Create";
            DateTime endTime = DateTime.Now;
            var deltaTime = endTime - startTime;
            var detlaMilliSec = deltaTime.Milliseconds;
            UIApplication uiApp = commandData.Application;
            HelperClassLibrary.logger.CreateDump(toolName, failedMessage, doc, uiApp, detlaMilliSec);
            return Result.Succeeded;
        }



    }

    class ColorScheme
    {
        public Color circulateC { get; }
        public Color meetC { get; }
        public Color operateC { get; }
        public Color washC { get; }
        public Color weC { get; }
        public Color workC { get; }
        public Color actualColor { get; set; }

        public ColorScheme(string programType)
        {
            this.circulateC = new Color(255, 247, 223);
            this.meetC = new Color(183, 240, 217);
            this.operateC = new Color(226, 226, 226);
            this.washC = new Color(195, 195, 195);
            this.weC = new Color(255, 210, 106);
            this.workC = new Color(171, 221, 231);
            switch(programType)
            {
                case "circulate":
                    this.actualColor = circulateC;
                    break;
                case "meet":
                    this.actualColor = meetC;
                    break;
                case "operate":
                    this.actualColor = operateC;
                    break;
                case "wash":
                    this.actualColor = washC;
                    break;
                case "we":
                    this.actualColor = weC;
                    break;
                case "work":
                    this.actualColor = workC;
                    break;
                default:
                    this.actualColor = washC;
                    break;
            }
        }
    }
}