using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Controls;
using Autodesk.Revit.Creation;
using Document = Autodesk.Revit.DB.Document;

namespace SMRevitTools.Classes.EnlargedDrawing
{
    public class EnlargedDrawingMethods
    {
        public static List<Element> GetSelection(UIDocument uidoc)
        {
            //filterout casework, furniture/system/ceilings from user selection and return a List of Elements
            List<Element> caseworkList = new List<Element>();
            Selection selection = uidoc.Selection;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            if (0 == selectedIds.Count())
            {
                // If no elements are selected.
                TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "You haven't selected anything!");
            }
            else
            {
                foreach (ElementId id in selectedIds)
                {
                    Element elem = uidoc.Document.GetElement(id);

                    ElementId groupId = elem.GroupId;
                    if (groupId != null)
                    {
                        caseworkList.Add(elem);
                    }
                    else if (elem.Category.Name.ToLower().Contains("casework")
                        || elem.Category.Name.ToLower().Contains("furniture")
                        || elem.Category.Name.ToLower().Contains("ceiling"))
                    {
                        caseworkList.Add(elem);
                    }
                }
                if (0 == caseworkList.Count())
                {
                    // If no furniture/casework are selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Casework/Furniture/Furniture System/Ceiling!");
                }
            }
            return caseworkList;
        }

        //Alternative Way for Bounding Box from List of elements
        public static BoundingBoxXYZ GetDirectBoundingBox(List<Element> caseworkList, Autodesk.Revit.DB.Document doc)
        {

            List<double> xCoords = new List<double>();
            List<double> yCoords = new List<double>();
            List<double> zCoords = new List<double>();

            List<XYZ> minPoints = new List<XYZ>();
            List<XYZ> maxPoints = new List<XYZ>();
            foreach (Element e in caseworkList)
            {
                BoundingBoxXYZ tempBB = e.get_BoundingBox(doc.ActiveView);
                xCoords.Add(tempBB.Min.X);
                yCoords.Add(tempBB.Min.Y);
                zCoords.Add(tempBB.Min.Z);

                xCoords.Add(tempBB.Max.X);
                yCoords.Add(tempBB.Max.Y);
                zCoords.Add(tempBB.Max.Z);
            }

            XYZ minima = new XYZ(xCoords.Min() - 0.5, yCoords.Min() - 0.5, zCoords.Min() - 0.5);
            XYZ maxima = new XYZ(xCoords.Max() + 0.5, yCoords.Max() + 0.5, zCoords.Max() + 0.5);

            BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
            offsetBoundingBox.Min = minima;
            offsetBoundingBox.Max = maxima;

            return offsetBoundingBox;
        }

        public static Solid GetUnionSolid(List<Element> caseworkList)
        {
            //Combined Solid for entire casework selection
            List<GeometryObject> geometryObjects = new List<GeometryObject>();
            Options options = new Options
            {
                ComputeReferences = true, // To compute references for accurate geometry
                //IncludeNonVisibleObjects = true // To include non-visible objects (e.g., hidden elements)
            };


            Solid unionSolid = null;
            List<Solid> solids = new List<Solid>();
            foreach (Element elem in caseworkList)
            {
                if (elem != null)
                {
                    GeometryElement geoElem = elem.get_Geometry(options);

                    foreach (GeometryObject geoObject in geoElem)
                    {

                        if (geoObject is Solid)
                        {
                            Solid solid = (Solid)geoObject;
                            if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                            {
                                solids.Add(solid);
                            }
                            // Single-level recursive check of instances. If viable solids are more than
                            // one level deep, this example ignores them.
                        }
                        else if (geoObject is GeometryInstance)
                        {
                            GeometryInstance geomInst = (GeometryInstance)geoObject;
                            GeometryElement instGeomElem = geomInst.GetInstanceGeometry();
                            foreach (GeometryObject instGeomObj in instGeomElem)
                            {
                                if (instGeomObj is Solid)
                                {
                                    Solid solid = (Solid)instGeomObj;
                                    if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                                    {
                                        solids.Add(solid);
                                    }
                                }
                            }
                        }
                    }
                }
            }


            foreach (Solid solid in solids)
            {
                if (solid != null)
                {
                    if (unionSolid == null)
                    {
                        unionSolid = solid;
                    }
                    else
                    {
                        unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);

                    }
                }
            }

            return unionSolid;
        }


        public static BoundingBoxXYZ GetBB(Solid unionSolid)
        {
            //Get bounidng box of Union solid with 0.5 feet offset
            BoundingBoxXYZ boundingBox = new BoundingBoxXYZ();
            BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();

            if (unionSolid != null)
            {
                // Calculate the bounding box
                boundingBox = unionSolid.GetBoundingBox();
                XYZ min = boundingBox.Min;
                XYZ max = boundingBox.Max;
                Autodesk.Revit.DB.Transform tForm = boundingBox.Transform;
                min = new XYZ(min.X - 0.5, min.Y - 0.5, min.Z - 0.5);
                max = new XYZ(max.X + 0.5, max.Y + 0.5, max.Z + 0.5);
                offsetBoundingBox.Min = min;
                offsetBoundingBox.Max = max;
                offsetBoundingBox.Transform = tForm;
            }

            return offsetBoundingBox;
        }

        public static string[] CheckSheetName(string[] sheetNameNumber)
        {
            string[] checkedString = sheetNameNumber;
            if (checkedString[0] == null)
            {
                checkedString[0] = "Unnamed";
            }
            //else if (checkedString[0].ToLower().Contains("dhruv"))
            //{
            //    checkedString[0] = "He who mustNotBeNamed";
            //}
            if (checkedString[1] == null || checkedString[1] == "00")
            {
                checkedString[1] = "Unknown floor";
                TaskDialog.Show("Invalid Input", "Invalid input : Floor not selected ???");
            }

            return checkedString;
        }

        public static string GetNextSheetNumber(Document doc, string geometricLevel)
        {
            //List of all sheets
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>().Where(x => x.SheetNumber.ToLower().Contains("b1-i-71"))
                .ToList();

            //Get Next Sheet Number in List
            int sheetIncreament = 01;
            string sheetNumber = "B1-I-71F" + geometricLevel + "-01";
            bool sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            while (sheetNumberExists)
            {
                sheetIncreament++;
                if(sheetIncreament <10)
                {
                    sheetNumber = "B1-I-71F" + geometricLevel + "-0" + sheetIncreament.ToString();
                    sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
                    continue;
                }
                sheetNumber = "B1-I-71F" + geometricLevel + "-" + sheetIncreament.ToString();
                sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            }
            return sheetNumber;
        }

        public static string GetNextSheetNumberIncreament(Document doc, string geometricLevel)
        {
            //List of all sheets
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>().Where(x => x.SheetNumber.ToLower().Contains("b1-i-30"))
                .ToList();

            //Get Next Sheet Number in List
            int sheetIncreament = 01;
            string sheetNumber = "B1-I-30" + "01";
            bool sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            while (sheetNumberExists)
            {
                sheetIncreament++;
                if (sheetIncreament < 10)
                {
                    sheetNumber = "B1-I-30" + "0" + sheetIncreament.ToString();
                    sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
                    continue;
                }
                sheetNumber = "B1-I-30" + sheetIncreament.ToString();
                sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            }
            return sheetNumber;
        }


        public static ElementId GetPackBTitleBlockId(Document doc)
        {
            FamilySymbol titleBlock = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .FirstOrDefault(q => q.Name == "A1") as FamilySymbol;
            ElementId titleBlockId = titleBlock.Id;
            //Get Titleblock Family 

            return titleBlockId;
        }

        public static List<BoundingBoxXYZ> GetViewSectionBB(XYZ maxCasework, XYZ minCasework, XYZ maxCalloutPoint, XYZ minCalloutPoint)
        {
            //Creating Section View bounding box
            double w = maxCasework.Y - minCasework.Y;
            double d = maxCasework.X - minCasework.X;
            XYZ sectionMin = new XYZ(-w * 0.5, 0, 0);
            XYZ sectionMax = new XYZ(w * 0.5, maxCasework.Z, d * 0.5);
            XYZ midPoint = (minCalloutPoint + maxCalloutPoint) * 0.5;
            XYZ secDir = XYZ.BasisY;
            XYZ up = XYZ.BasisZ;
            XYZ viewdir = secDir.CrossProduct(up);

            //Transform object for bounding box
            Transform t = Transform.Identity;
            Transform t2 = t;

            t.Origin = midPoint;
            t.BasisX = secDir;
            t.BasisY = up;
            t.BasisZ = viewdir;
            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
            sectionBox.Transform = t;
            sectionBox.Min = sectionMin;
            sectionBox.Max = sectionMax;


            XYZ secDir2 = XYZ.BasisX;
            XYZ viewdir2 = secDir2.CrossProduct(up);
            double w2 = maxCasework.Y - minCasework.Y;
            double d2 = maxCasework.X - minCasework.X;
            XYZ sectionMin2 = new XYZ(-d2 * 0.5, 0, 0);
            XYZ sectionMax2 = new XYZ(d2 * 0.5, maxCasework.Z, w2 * 0.5);
            t2.Origin = midPoint;
            t2.BasisX = secDir2;
            t2.BasisY = up;
            t2.BasisZ = viewdir2;
            BoundingBoxXYZ sectionBox2 = new BoundingBoxXYZ();
            sectionBox2.Transform = t2;
            sectionBox2.Min = sectionMin2;
            sectionBox2.Max = sectionMax2;

            List<BoundingBoxXYZ> sectionBoxes = new List<BoundingBoxXYZ>();
            sectionBoxes.Add(sectionBox);
            sectionBoxes.Add(sectionBox2);
            return sectionBoxes;
        }

    }
}
