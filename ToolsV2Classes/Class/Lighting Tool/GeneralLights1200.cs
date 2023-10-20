using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class GeneralLights1200 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;

            try
            {
                //Get Lighting fixture Type for Office Space
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                FamilySymbol light2400Symbol = collector.OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .First(x => x.Name == "WWI-LT-LS03-03");
                FamilySymbol light1200Symbol = collector.OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .First(x => x.Name == "WWI-LT-LS03-01");
                FamilySymbol lightCanSymbol = collector.OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .First(x => x.Name == "WWI-LT-LS02-01");

                ////Pick Room
                //TaskDialog.Show("Pick Room", "Select the Room Element");
                //Reference pickedRoom = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                ////Retrive Element ID
                //ElementId roomId = pickedRoom.ElementId;
                //Room roomtest = doc.GetElement(roomId) as Room;
                //string roomName = roomtest.Name;

                string programType;
                View activeView = doc.ActiveView;
                FilteredElementCollector collector2 = new FilteredElementCollector(doc, activeView.Id).OfClass(typeof(SpatialElement));
                List<Room> rooms = new List<Room>();
                foreach (SpatialElement e in collector2)
                {
                    programType = GetParamVal(e, "WW-ProgramType");
                    if (programType == "Work" || programType == "We")
                    {
                        rooms.Add(e as Room);
                    }
                }
                using (Transaction trans = new Transaction(doc, "Create Lighting Fixtures"))
                {
                    trans.Start();
                    int noOfRooms = rooms.Count;
                    if (noOfRooms > 0)
                    {
                        foreach (Room room in rooms)
                        {
                            //Get Room Bounding box for simpler rooms
                            BoundingBoxXYZ roomBox = room.get_BoundingBox(null);
                            programType = GetParamVal(room, "WW-ProgramType");
                            if (programType == "Work")
                            {
                                ModelOfficeLights(light1200Symbol, roomBox, room, doc);
                            }
                            else if (programType == "We")
                            {
                                ModelLoungeLights(lightCanSymbol, roomBox, room, doc);
                            }
                        }
                    }
                    else
                    {
                        TaskDialog.Show("View Error", "Any of the rooms are not visible in view, please update the visibility settings and try again");

                    }

                    trans.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }

        public void ModelOfficeLights(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc)
        {
            if (!lightSymbol.IsActive)
            {
                lightSymbol.Activate();
            }
            if (roomBox != null)
            {
                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;

                double L = xMaxVal - xMinVal;
                double W = yMaxVal - yMinVal;

                if ( L >= W )
                {
                    int nMin = (int)Math.Ceiling(L / 2700);
                    int nMax = (int)Math.Floor(L / 2100);

                    int mMin = (int)Math.Ceiling(W / 3600);
                    int mMax = (int)Math.Ceiling((W - 600) / 3000);


                    int possibleLayouts = 0;
                    List<int> nList = new List<int>();
                    for (int i = nMin; i <= nMax; i++)
                    {
                        nList.Add(i);
                        possibleLayouts++;
                    }
                    double x = L / (2 * nMin);

                    for (int j = 1; j <= mMax; j++)
                    {
                        for (int i = 1; i <= nMin; i++)
                        {
                            XYZ tempPoint1 = new XYZ((xMinVal + x + (i - 1) * 2 * x) / 304.8, (yMinVal + 1800 + (j - 1) * 3000) / 304.8, zVal / 304.8);
                            XYZ tempPoint2 = new XYZ((xMinVal + x + (i - 1) * 2 * x) / 304.8, (yMinVal + 1800 + (j - 1) * 3000) / 304.8, (zVal + 1000) / 304.8);
                            bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                            if (lightInsideRoom)
                            {
                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                Line axis2 = Line.CreateBound(tempPoint1, tempPoint2);
                                ElementTransformUtils.RotateElement(doc, fi2.Id, axis2, Math.PI / 2);
                                SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            }
                        }
                    }
                }
                else
                {
                    int mMin = (int)Math.Ceiling(W / 2700);
                    int mMax = (int)Math.Floor(W / 2100);

                    int nMin = (int)Math.Ceiling(L / 3600);
                    int nMax = (int)Math.Ceiling((L - 600) / 3000);


                    int possibleLayouts = 0;
                    List<int> nList = new List<int>();
                    for (int i = nMin; i <= nMax; i++)
                    {
                        nList.Add(i);
                        possibleLayouts++;
                    }
                    double x = W / (2 * mMin);

                    for (int j = 1; j <= nMax; j++)
                    {
                        for (int i = 1; i <= mMin; i++)
                        {
                            XYZ tempPoint1 = new XYZ((xMinVal + 1800 + (j - 1) * 3000) / 304.8, (yMinVal + x + (i - 1) * 2 * x) / 304.8, zVal / 304.8);
                            XYZ tempPoint2 = new XYZ((xMinVal + 1800 + (j - 1) * 3000) / 304.8, (yMinVal + x + (i - 1) * 2 * x) / 304.8, (zVal + 1000) / 304.8);
                            bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                            if (lightInsideRoom)
                            {
                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                            }
                        }
                    }
                }
                

            }
        }

        public void ModelLoungeLights(FamilySymbol lightSymbol, BoundingBoxXYZ roomBox, Room room1, Document doc)
        {
            if (!lightSymbol.IsActive)
            {
                lightSymbol.Activate();
            }
            if (roomBox != null)
            {
                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;

                double L = (xMaxVal - xMinVal);
                double W = (yMaxVal - yMinVal);


                int nMin = (int)Math.Ceiling((L - 2400 + 2200) / 2200);
                int nMax = (int)Math.Floor((L - 2400 + 1500) / 1500);

                int mMin = (int)Math.Ceiling((W - 2400 + 2200) / 2200);
                int mMax = (int)Math.Floor((W - 2400 + 1500) / 1500);


                double x = (L - 2400) / (nMin - 1);
                double y = (W - 2400) / (mMin - 1);

                for (int j = 1; j <= mMin; j++)
                {
                    for (int i = 1; i <= nMin; i++)
                    {
                        XYZ tempPoint1 = new XYZ((xMinVal + 1200 + (i - 1) * x) / 304.8, (yMinVal + 1200 + (j - 1) * y) / 304.8, zVal / 304.8);
                        XYZ tempPoint2 = new XYZ((xMinVal + 1200 + (i - 1) * x) / 304.8, (yMinVal + 1200 + (j - 1) * y) / 304.8, (zVal + 1000) / 304.8);
                        bool lightInsideRoom = room1.IsPointInRoom(tempPoint2);
                        if (lightInsideRoom)
                        {
                            FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            Line axis2 = Line.CreateBound(tempPoint1, tempPoint2);
                            ElementTransformUtils.RotateElement(doc, fi2.Id, axis2, Math.PI / 2);
                            SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                        }
                    }
                }

            }
        }

        public String GetParamVal(SpatialElement r, string sharedParameter)
        {
            String paraValue;
            Guid paraGuid = r.LookupParameter(sharedParameter).GUID;
            paraValue = r.get_Parameter(paraGuid).AsString();

            return paraValue;
        }
        public void SetBOLFVal(FamilyInstance fi, string sharedParameter, double BOFH)
        {
            Parameter para = fi.LookupParameter(sharedParameter);
            para.Set(BOFH/ 304.8);
        }


        //public static List<XYZ> ConvexHull(List<XYZ> points)
        //{
        //    if (points == null) throw new ArgumentNullException(nameof(points));
        //    XYZ startPoint = points.MinBy(p => p.X);
        //    var convexHullPoints = new List<XYZ>();
        //    XYZ walkingPoint = startPoint;
        //    XYZ refVector = XYZ.BasisY.Negate();
        //    do
        //    {
        //        convexHullPoints.Add(walkingPoint);
        //        XYZ wp = walkingPoint;
        //        XYZ rv = refVector;
        //        walkingPoint = points.MinBy(p =>
        //        {
        //            double angle = (p - wp).AngleOnPlaneTo(rv, XYZ.BasisZ);
        //            if (angle < 1e-10) angle = 2 * Math.PI;
        //            return angle;
        //        });
        //        refVector = wp - walkingPoint;
        //    } while (walkingPoint != startPoint);
        //    convexHullPoints.Reverse();
        //    return convexHullPoints;
        //}

        //static List<XYZ> GetConvexHullOfRoomBoundary(IList<IList<BoundarySegment>> boundary)
        //{
        //    List<XYZ> pts = new List<XYZ>();

        //    foreach (IList<BoundarySegment> loop in boundary)
        //    {
        //        foreach (BoundarySegment seg in loop)
        //        {
        //            Curve c = seg.GetCurve();
        //            pts.AddRange(c.Tessellate());
        //        }
        //    }
        //    int n = pts.Count;

        //    return GetLitIndia.ConvexHull(pts);
        //}

    }

    //public static class IEnumerableExtensions
    //{
    //    public static tsource MinBy<tsource, tkey>(
    //      this IEnumerable<tsource> source,
    //      Func<tsource, tkey> selector)
    //    {
    //        return source.MinBy(selector, Comparer<tkey>.Default);
    //    }
    //    public static tsource MinBy<tsource, tkey>(
    //      this IEnumerable<tsource> source,
    //      Func<tsource, tkey> selector,
    //      IComparer<tkey> comparer)
    //    {
    //        if (source == null) throw new ArgumentNullException(nameof(source));
    //        if (selector == null) throw new ArgumentNullException(nameof(selector));
    //        if (comparer == null) throw new ArgumentNullException(nameof(comparer));
    //        using (IEnumerator<tsource> sourceIterator = source.GetEnumerator())
    //        {
    //            if (!sourceIterator.MoveNext())
    //                throw new InvalidOperationException("Sequence was empty");
    //            tsource min = sourceIterator.Current;
    //            tkey minKey = selector(min);
    //            while (sourceIterator.MoveNext())
    //            {
    //                tsource candidate = sourceIterator.Current;
    //                tkey candidateProjected = selector(candidate);
    //                if (comparer.Compare(candidateProjected, minKey) < 0)
    //                {
    //                    min = candidate;
    //                    minKey = candidateProjected;
    //                }
    //            }
    //            return min;
    //        }
    //    }
    //}

}
