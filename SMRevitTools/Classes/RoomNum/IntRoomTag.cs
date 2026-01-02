using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;

namespace SMRevitTools
{
    [Transaction(TransactionMode.Manual)]
    public class IntRoomTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            try
            {
                // 1. Gather all ROOM TAGS in the current view
                var allTags = new FilteredElementCollector(doc, activeView.Id)
                    .OfClass(typeof(SpatialElementTag))
                    .OfType<SpatialElementTag>()
                    .Where(t => t.Category.Id.Value == (int)BuiltInCategory.OST_RoomTags)
                    .ToList();

                // Find host model room IDs already tagged
                var alreadyTaggedHostRoomIds = new HashSet<ElementId>();
                // Find linked model room keys already tagged: (RevitLinkInstance Id, Room Id)
                var alreadyTaggedLinkedRooms = new HashSet<(ElementId linkInstId, ElementId linkedRoomId)>();

                foreach (var tag in allTags)
                {
                    ElementId taggedId = tag.Id;
                    if (taggedId != ElementId.InvalidElementId)
                    {
                        Element taggedElem = (tag as RoomTag).Room;
                        if (taggedElem != null && taggedElem is Room)
                        {
                            // host model room tag
                            alreadyTaggedHostRoomIds.Add(taggedElem.Id);
                        }
                        else if (taggedElem == null)
                        {
                            // Check if this is a room in a linked model
                            // IndependentTag.TaggedLocalElementId gives the local (host or linked) room id
                            // IndependentTag.TaggedLinkedElementId gives the link room id (if tag is for a link)

                            var taggedRoomLinkId = (tag as RoomTag).TaggedRoomId;
                            var taggedRoomId = taggedRoomLinkId.LinkedElementId;
                            
                            if (taggedRoomId != ElementId.InvalidElementId)
                            {
                                alreadyTaggedLinkedRooms.Add((taggedRoomLinkId.LinkInstanceId, taggedRoomLinkId.LinkedElementId));
                            }
                        }
                    }
                }

                // 2. Gather all host model rooms in the view
                var roomsHost = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .OfType<Room>()
                    .Where(r => r.Area > 0 && r.Location is LocationPoint)
                    .ToList();

                // 3. Gather all rooms from linked models visible in the view
                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .OfType<RevitLinkInstance>()
                    .ToList();

                var roomsFromLinks = new List<(Room room, RevitLinkInstance instance, XYZ pointHostCoords)>();

                foreach (var linkInst in linkInstances)
                {
                    Document linkDoc = linkInst.GetLinkDocument();
                    if (linkDoc == null) continue;

                    // Get all rooms in the link
                    var linkRooms = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType()
                        .OfType<Room>();

                    foreach (var linkRoom in linkRooms)
                    {
                        if (linkRoom.Area <= 0 || !(linkRoom.Location is LocationPoint locPt)) continue;

                        // The room's origin in the link's coordinate system:
                        locPt = linkRoom.Location as LocationPoint;
                        XYZ linkPoint = locPt.Point;
                        // Transform to host coordinates:
                        Transform linkTransform = linkInst.GetTransform();
                        XYZ hostPoint = linkTransform.OfPoint(linkPoint);

                        // Optional: Only tag if the point falls within view's bounding box (view crop box),
                        // Uncomment below section for filtering per view
                        /*
                        BoundingBoxXYZ viewBbox = activeView.CropBox;
                        if (viewBbox != null && !IsPointWithinBoundingBox(hostPoint, viewBbox))
                            continue;
                        */

                        roomsFromLinks.Add((linkRoom, linkInst, hostPoint));
                    }
                }

                using (Transaction tx = new Transaction(doc, "Tag All Untagged Rooms in View"))
                {
                    tx.Start();

                    // 4. Tag untagged host rooms
                    foreach (var room in roomsHost)
                    {
                        if (alreadyTaggedHostRoomIds.Contains(room.Id)) continue;
                        var locPt = room.Location as LocationPoint;
                        if (locPt == null) continue;
                        UV tagUV = new UV(locPt.Point.X, locPt.Point.Y);

                        doc.Create.NewRoomTag(new LinkElementId(room.Id), tagUV, activeView.Id);
                    }

                    // 5. Tag untagged linked rooms

                    FamilySymbol roomTagSymbol = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_RoomTags)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault();

                    if (roomTagSymbol == null)
                        throw new InvalidOperationException("No room tag family symbol found.");
                    foreach (var (linkRoom, linkInstance, hostPoint) in roomsFromLinks)
                    {
                        var key = (linkInstance.Id, linkRoom.Id);
                        if (alreadyTaggedLinkedRooms.Contains(key)) continue;

                        UV tagUV = new UV(hostPoint.X, hostPoint.Y);

                        BoundingBoxXYZ cropBox = activeView.CropBox;
                        if (hostPoint.X < cropBox.Min.X || hostPoint.X > cropBox.Max.X ||
                            hostPoint.Y < cropBox.Min.Y || hostPoint.Y > cropBox.Max.Y)
                        {
                            // Skip if point is outside crop box
                            continue;
                        }

                        // Use NewRoomTag for linked rooms: this creates a valid tag without needing a Reference
                        doc.Create.NewRoomTag(new LinkElementId(linkInstance.Id, linkRoom.Id), tagUV, activeView.Id);
                    }

                    tx.Commit();
                }

                TaskDialog.Show("Tag Rooms", "All untagged rooms in this view (host and linked) are now tagged.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "Room Tag Tool: " + ex.ToString();
                return Result.Failed;
            }
        }

        // (Optional utility for view crop box check)
        private static bool IsPointWithinBoundingBox(XYZ pt, BoundingBoxXYZ bbox)
        {
            return (pt.X >= bbox.Min.X && pt.X <= bbox.Max.X &&
                    pt.Y >= bbox.Min.Y && pt.Y <= bbox.Max.Y &&
                    pt.Z >= bbox.Min.Z && pt.Z <= bbox.Max.Z);
        }
    }
}
