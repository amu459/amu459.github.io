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
    public class GetLitIndia : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;

            try
            {
                //Pick Room
                TaskDialog.Show("Pick Room", "Select the Room Element");
                Reference pickedRoom = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                //Retrive Element ID
                ElementId roomId = pickedRoom.ElementId;
                Room room = doc.GetElement(roomId) as Room;
                string roomName = room.Name;
                string programType = GetParamVal(room, "WW-ProgramType");





                //Get Room Bounding box for simpler rooms
                BoundingBoxXYZ roomBox = room.get_BoundingBox(null);
                XYZ roomMin = roomBox.Min;
                XYZ roomMax = roomBox.Max;
                double zVal = roomMin.Z * 304.8;
                double xMinVal = roomMin.X * 304.8;
                double yMinVal = roomMin.Y * 304.8;
                double xMaxVal = roomMax.X * 304.8;
                double yMaxVal = roomMax.Y * 304.8;

                ////Get Convex Hull of Room
                //SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
                //opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                //IList<IList<BoundarySegment>> boundarySegmentArray = room.GetBoundarySegments(opt);
                //List<XYZ> roomConvexHullPoints = GetConvexHullOfRoomBoundary(boundarySegmentArray);

                double L = xMaxVal - xMinVal;
                double W = yMaxVal - yMinVal;
                string isHorizontal = "L";
                if (L < W)
                {
                    double tempLength = L;
                    //double tempx = xMinVal;
                    //xMinVal = yMinVal;
                    //yMinVal = tempx;

                    //tempx = xMaxVal;
                    //xMaxVal = yMaxVal;
                    //yMaxVal = tempx;

                    L = W;
                    W = tempLength;
                    isHorizontal = "W";
                }

                int nMin = (int)Math.Ceiling(L / 2700);
                int nMax = (int)Math.Floor(L / 2100);

                int mMin = (int)Math.Ceiling(W / 3600);
                int mMax = (int)Math.Ceiling((W - 600) / 3000);


                int possibleLayouts = 0;
                List<int> nList = new List<int>();
                for (int i=nMin; i<=nMax; i++)
                {
                    nList.Add(i);
                    possibleLayouts++;
                }


                using (Transaction trans = new Transaction(doc, "Create Lighting Fixtures"))
                {
                    trans.Start();
                    //Get Lighting fixture Type for Office Space
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    FamilySymbol lightSymbol = collector.OfClass(typeof(FamilySymbol))
                        .WhereElementIsElementType()
                        .Cast<FamilySymbol>()
                        .First(x => x.Name == "L22-24");

                    string lightTypeName = lightSymbol.Name;
                    if (!lightSymbol.IsActive)
                    {
                        lightSymbol.Activate();
                    }

                    double x = L / (2 * nMin);

                    if (isHorizontal == "L")
                    {

                        for (int j = 1; j <= mMax; j++)
                        {

                            for (int i = 1; i <= nMin; i++)
                            {
                                XYZ tempPoint1 = new XYZ((xMinVal + x + (i - 1) * 2 * x) / 304.8, (yMinVal + 1800 + (j-1)*3000) / 304.8, zVal / 304.8);
                                XYZ tempPoint2 = new XYZ((xMinVal + x + (i - 1) * 2 * x) / 304.8, (yMinVal + 1800 + (j - 1) * 3000) / 304.8, (zVal + 1000) / 304.8);
                                bool lightInsideRoom = room.IsPointInRoom(tempPoint2);
                                if (lightInsideRoom )
                                {
                                    FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    Line axis2 = Line.CreateBound(tempPoint1, tempPoint2);
                                    ElementTransformUtils.RotateElement(doc, fi2.Id, axis2, Math.PI / 2);
                                    SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                                }
                            }
                        }
                    }
                    else if (isHorizontal == "W")
                    {

                        for (int j = 1; j <= mMax; j++)
                        {

                            for (int i = 1; i <= nMin; i++)
                            {
                                XYZ tempPoint1 = new XYZ((xMinVal + 1800 + (j - 1) * 3000) / 304.8, (yMinVal + x + (i - 1) * 2 * x) / 304.8, zVal / 304.8);
                                XYZ tempPoint2 = new XYZ((xMinVal + 1800 + (j - 1) * 3000) / 304.8, (yMinVal + x + (i - 1) * 2 * x) / 304.8, (zVal + 1000) / 304.8);
                                bool lightInsideRoom = room.IsPointInRoom(tempPoint2);
                                if(lightInsideRoom)
                                {
                                    FamilyInstance fi2 = doc.Create.NewFamilyInstance(tempPoint1, lightSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    SetBOLFVal(fi2, "WW-BottomOfFixtureHeight", 2400);
                                }
                            }
                        }
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
