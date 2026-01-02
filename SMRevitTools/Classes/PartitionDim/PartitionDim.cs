using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class PartitionDim : IExternalCommand
    {

        // Implement the Execute method
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            try
            {
                // Create a filtered element collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);


                View activeView = doc.ActiveView;

                // Prepare a 3D view for ReferenceIntersector
                View3D view3d = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D)).Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);
                if (view3d == null)
                {
                    TaskDialog.Show("Error", "No suitable 3D view found.");
                    return Result.Failed;
                }

                // Get selected rooms
                var selectedIds = uidoc.Selection.GetElementIds();
                var rooms = selectedIds.Select(id => doc.GetElement(id)).OfType<Room>().ToList();
                if (rooms.Count == 0)
                {
                    TaskDialog.Show("Error", "Please select one or more rooms.");
                    return Result.Failed;
                }

                // Find the DimensionType
                var dimType = new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType))
                    .Cast<DimensionType>()
                    .FirstOrDefault(dt => dt.Name.Equals("SM_Rounded_2.0mm", StringComparison.OrdinalIgnoreCase));
                if (dimType == null)
                {
                    TaskDialog.Show("Error", "Dimension style 'SM_Rounded_2.0mm' not found.");
                    return Result.Failed;
                }





                // Create rooms to specific family locations
                using (Transaction transaction = new Transaction(doc, "Tool Name"))
                {
                    transaction.Start();


                    foreach (Room room in rooms)
                    {
                        XYZ roomCenter = GetRoomCenter(room);
                        if (roomCenter == null) continue;

                        // Get wall IDs that bound this room (including subwalls for stacked wall support)
                        var boundaryWallIds = GetBoundaryWallAndSubWallIds(room);

                        // Four main rays for axes
                        var directions = new Dictionary<string, XYZ>
                {
                    {"X+", new XYZ(1, 0, 0)},
                    {"X-", new XYZ(-1, 0, 0)},
                    {"Y+", new XYZ(0, 1, 0)},
                    {"Y-", new XYZ(0, -1, 0)}
                };

                        var axisRefs = new Dictionary<string, (Reference Ref, XYZ HitPoint)>();

                        foreach (var pair in directions)
                        {
                            string key = pair.Key;
                            XYZ dir = pair.Value;
                            double epsilon = 0.01; // 3mm fudge factor inside the room
                            XYZ localOrigin = roomCenter + dir * epsilon;

                            ReferenceIntersector intersector = new ReferenceIntersector(
                                new ElementCategoryFilter(BuiltInCategory.OST_Walls),
                                FindReferenceTarget.Face,
                                view3d);

                            var hits = intersector.Find(localOrigin, dir)
                                .Where(rwc => {
                                    Element wallElem = doc.GetElement(rwc.GetReference().ElementId);
                                    return wallElem is Wall && !wallElem.Document.IsLinked && boundaryWallIds.Contains(wallElem.Id);
                                })
                                .OrderBy(rwc => rwc.Proximity)
                                .ToList();

                            if (hits.Count > 0)
                            {
                                Wall wall = doc.GetElement(hits[0].GetReference().ElementId) as Wall;
                                if (wall != null)
                                {
                                    // KEY: Use face closest to actual intersection and roughly facing room center
                                    Reference bestRef = FindBestRoomFaceReference(wall, hits[0].GetReference().GlobalPoint, roomCenter);
                                    if (bestRef != null)
                                        axisRefs[key] = (bestRef, hits[0].GetReference().GlobalPoint);
                                }
                            }
                        }

                        // Compose ReferenceArray and create dimensions for each axis, offset for clarity
                        if (axisRefs.ContainsKey("X+") && axisRefs.ContainsKey("X-"))
                            CreateDimensionOnView(doc, uidoc.ActiveView, dimType, axisRefs["X-"], axisRefs["X+"], roomCenter, true);
                        if (axisRefs.ContainsKey("Y+") && axisRefs.ContainsKey("Y-"))
                            CreateDimensionOnView(doc, uidoc.ActiveView, dimType, axisRefs["Y-"], axisRefs["Y+"], roomCenter, false);
                    }

                    transaction.Commit();
                }

                // Return success result

                string toolName = "Tool Name";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success - ", doc, uiApp, detlaMilliSec);
            skipped:
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Tool Name";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }





        // Helper: Find wall face closest to intersection and pointing into room (robust for all wall types)
        private Reference FindBestRoomFaceReference(Wall wall, XYZ intersectionPoint, XYZ roomCenter)
        {
            Options opt = new Options { ComputeReferences = true };
            GeometryElement geomElem = wall.get_Geometry(opt);

            Reference bestRef = null;
            double bestDist = double.MaxValue;

            foreach (GeometryObject go in geomElem)
            {
                Solid solid = go as Solid;
                if (solid == null) continue;

                foreach (Face face in solid.Faces)
                {
                    PlanarFace pf = face as PlanarFace;
                    if (pf == null) continue;

                    double dist = pf.Origin.DistanceTo(intersectionPoint);
                    // Inwardness: at least moderately facing toward room center
                    XYZ toRoom = (roomCenter - pf.Origin).Normalize();
                    double dot = pf.FaceNormal.Normalize().DotProduct(toRoom);

                    // Must "face into" room, or at least not point away
                    if (dot > 0.3 && dist < bestDist)
                    {
                        bestRef = pf.Reference;
                        bestDist = dist;
                    }
                }
            }
            return bestRef;
        }

        // Helper: add all bounding wall and subwall IDs for robust stacked wall support
        private HashSet<ElementId> GetBoundaryWallAndSubWallIds(Room room)
        {
            var ids = new HashSet<ElementId>();
            var opts = new SpatialElementBoundaryOptions();
            var loops = room.GetBoundarySegments(opts);
            if (loops == null) return ids;

            foreach (var loop in loops)
                foreach (var seg in loop)
                {
                    if (seg.ElementId.IntegerValue == -1) continue;
                    var elem = room.Document.GetElement(seg.ElementId);
                    if (elem is Wall wall)
                    {
                        if (wall.IsStackedWall)
                        {
                            foreach (ElementId subWallId in wall.GetStackedWallMemberIds())
                                ids.Add(subWallId);
                        }
                        else
                        {
                            ids.Add(wall.Id);
                        }
                    }
                }
            return ids;
        }

        // Helper: get center of room from LocationPoint or bounding box
        private XYZ GetRoomCenter(Room room)
        {
            var loc = room.Location as LocationPoint;
            if (loc != null) return loc.Point;
            var bb = room.get_BoundingBox(null);
            return bb != null ? (bb.Min + bb.Max) / 2 : null;
        }

        // Create dimension slightly offset outside room area for clarity
        private void CreateDimensionOnView(Document doc, View view, DimensionType dimType,
            (Reference Ref, XYZ Pt) neg, (Reference Ref, XYZ Pt) pos, XYZ roomCenter, bool isXAxis)
        {
            ReferenceArray refArray = new ReferenceArray();
            refArray.Append(neg.Ref);
            refArray.Append(pos.Ref);

            XYZ wallVec = (pos.Pt - neg.Pt).Normalize();
            XYZ offsetDir = isXAxis ? new XYZ(0, 1, 0) : new XYZ(1, 0, 0);
            double offsetDist = 1.5; // feet; adjust as needed

            XYZ mid = (neg.Pt + pos.Pt) / 2.0 + offsetDir * offsetDist;
            Line dimLine = Line.CreateBound(
                mid - wallVec * 10,
                mid + wallVec * 10
            );

            Dimension dim = doc.Create.NewDimension(view, dimLine, refArray);
            if (dim != null) dim.DimensionType = dimType;
        }
    }
}


