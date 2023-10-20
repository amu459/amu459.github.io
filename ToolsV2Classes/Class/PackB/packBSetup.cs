using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI.Selection;
using ToolsV2Classes.Class.PackB;
using Newtonsoft.Json.Linq;
using Autodesk.Revit.Creation;
using System.Windows.Shapes;
using Microsoft.Win32.SafeHandles;
using static System.Windows.Forms.LinkLabel;
using System.Net;
using Microsoft.Office.Interop.Excel;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class packBSetup : IExternalCommand
    {

        // Implement the Execute method
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;

            try
            {

                //Collect elements from user selection, filter out unrequired elements
                #region COLLECT ELEMENTS FROM SELECTION

                List<Element> caseworkList = new List<Element>();
                caseworkList = packBmethods.GetSelection(uidoc);
                if (0 == caseworkList.Count())
                {
                    // If no casework/furniture are selected.
                    goto skipTool;
                }
                int selectedItemsCount = caseworkList.Count();

                #endregion



                //List of Levels in the Project
                #region GET LIST OF LEVELS FROM REVIT

                FilteredElementCollector levelCollector = new FilteredElementCollector(doc).OfClass(typeof(Level));

                List<Level> levels = levelCollector
                    .Cast<Level>()
                    .Where(x => !x.Name.ToLower().Contains("container")).ToList();

                int levelCount = levels.Count();
                string[] levelNames = new string[levelCount];
                int count = 0;
                foreach (Level level in levels) 
                {
                    levelNames.SetValue(level.Name, count);
                    count++;
                }

                #endregion



                //Ask for Casework Name and select Floor
                #region USER INPUT FOR CASEWORK NAME AND LEVEL

                packBInputWindow inputWindow = new packBInputWindow(uidoc, levelNames);
                inputWindow.label_Count.Content = selectedItemsCount.ToString();
                inputWindow.ShowDialog();

                string[] sheetNameNumber = new string[2];

                //Sheet Name is set from Casework Name
                //Level Name as per selected Level from dropdown
                sheetNameNumber[0] = inputWindow.inputText;
                sheetNameNumber[1] = inputWindow.inputLevelName;
                if (sheetNameNumber[0] == "cancel")
                {
                    goto skipTool;
                }
                sheetNameNumber = packBmethods.CheckSheetName(sheetNameNumber);
                if (sheetNameNumber[1] == "Unknown floor")
                {
                    goto skipTool;
                }
                #endregion



                //Combined Solid for entire casework selection
                #region GET BOUNDING BOX OF SELECTION

                //Solid unionSolid = null;
                //unionSolid = packBmethods.GetUnionSolid(caseworkList);
                //if (unionSolid == null)
                //{
                //    // If union solid is null
                //    goto skipTool;
                //}


                //Get bounidng box of Union solid with 0.5 feet offset
                //BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
                //offsetBoundingBox = packBmethods.GetBB(unionSolid);
                //if (offsetBoundingBox == null)
                //{
                //    // If bounding box is null
                //    goto skipTool;
                //}

                BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
                offsetBoundingBox = packBmethods.GetDirectBoundingBox(caseworkList, doc);

                #endregion

                XYZ minCasework = offsetBoundingBox.Min;
                XYZ maxCasework = offsetBoundingBox.Max;



                //Required Sheet metadata
                #region SHEET METADATA

                string wwSheetCategory = "Package B"; //Default and unchagned
                string wwSheetSubCategory = "05 Interiors"; //Default and unchagned
                string wwSheetSer = "0001 Project Details"; //Changes as per Level selected
                string wwSheetIss = "•";
                if (sheetNameNumber[1] != "00")
                {
                    wwSheetSer = "0001 " + sheetNameNumber[1];
                }
                
                //Get list of sheet parameters
                ViewSheet defaultSheet = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>().FirstOrDefault();
                Guid sheetCategoryGuid = defaultSheet.LookupParameter("WW-SheetCategory").GUID;
                Guid sheetSubCategoryGuid = defaultSheet.LookupParameter("WW-SheetSubCategory").GUID;
                Guid sheetSeriesGuid = defaultSheet.LookupParameter("WW-SheetSeries").GUID;
                Guid sheetIssuancePackBGuid= defaultSheet.LookupParameter("WW-Sheet Issuance (Package B)").GUID;

                //Calculate next Sheet number based on sheet list
                string sheetNumber = packBmethods.GetNextSheetNumber(doc);

                //Get Titleblock Family for package BId
                ElementId packBTitleBlockId = packBmethods.GetPackBTitleBlockId(doc);

                #endregion



                //Retrive 3D View type
                #region RETRIVE 3D VIEW TYPE
                FilteredElementCollector viewTypeCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));

                ViewFamilyType desViewFamilyType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("packb3d")).FirstOrDefault();
                #endregion



                //Retrive Callout Plan view
                //Check which layout plan to create callout in
                #region RETRIVE CALLOUT VIEW TYPE
                Level layoutLevel = null;
                foreach(Level level in levels)
                {
                    if (level.Name.Equals(sheetNameNumber[1]))
                    {
                        layoutLevel = level;
                    }
                }
                if(layoutLevel == null)
                {
                    TaskDialog.Show("Null Error", "No Layout plan match found");
                    goto skipTool;
                }
                FilteredElementCollector viewCollector = new FilteredElementCollector(doc).OfClass(typeof(View));
                View layoutView = viewCollector.Cast<View>()
                    .Where(x => x.Name.Contains(sheetNameNumber[1])
                    && x.Name.Contains("Layout"))
                    .FirstOrDefault();
                ElementId layoutId = layoutView.Id;
                ViewFamilyType calloutType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("packbplan")).FirstOrDefault();
                ElementId calloutTypeId = calloutType.Id;
                #endregion
                //Get min and max point for callout views
                XYZ minCalloutPoint = new XYZ(minCasework.X, minCasework.Y, minCasework.Z);
                XYZ maxCalloutPoint = new XYZ(maxCasework.X, maxCasework.Y, minCasework.Z);



                //Retrive Key Plan type
                //Check which layout plan to create callout in
                #region RETRIVE KEYPLAN TYPE
                ViewFamilyType keyPlanType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("packbkeyplan")).FirstOrDefault();
                ElementId keyPlanTypeId = keyPlanType.Id;

                View keyPlanTemplate = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.ToLower().Contains("packbkeyplan")).FirstOrDefault();
                ElementId keyPlanTemplateId = keyPlanTemplate.Id;
                #endregion


                //Red Lines for Key Plan
                #region SETUP REDLINES ON KEYPLAN
                var gstyles = (new FilteredElementCollector(doc)).OfClass(typeof(GraphicsStyle)).Cast<GraphicsStyle>().ToList();
                GraphicsStyle _gstyle = gstyles.Where(x => x.GraphicsStyleType == GraphicsStyleType.Projection).FirstOrDefault(x => x.Name.ToLower().Contains("solid_red - key plan"));

                XYZ keyPoint1 = minCalloutPoint;
                XYZ keyPoint2 = new XYZ(maxCalloutPoint.X, minCalloutPoint.Y, minCalloutPoint.Z);
                XYZ keyPoint3 = maxCalloutPoint;
                XYZ keyPoint4 = new XYZ(minCalloutPoint.X, maxCalloutPoint.Y, minCalloutPoint.Z);

                Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(keyPoint1, keyPoint2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(keyPoint2, keyPoint3);
                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(keyPoint3, keyPoint4);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(keyPoint4, keyPoint1);
                #endregion



                //Create Section/Elevation
                #region RETRIVE SECTION TYPE
                ViewFamilyType sectionType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("packbsection")).FirstOrDefault();
                ElementId sectionTypeId = sectionType.Id;

                List<BoundingBoxXYZ> sectionBoxes = packBmethods.GetViewSectionBB(maxCasework, minCasework, maxCalloutPoint, minCalloutPoint);
                BoundingBoxXYZ sectionBox = sectionBoxes[0];
                BoundingBoxXYZ sectionBox2 = sectionBoxes[1];



                ////Creating Section View bounding box
                //double w = maxCasework.Y - minCasework.Y;
                //double d = maxCasework.X - minCasework.X;
                //XYZ sectionMin = new XYZ(-w * 0.5, 0, 0);
                //XYZ sectionMax = new XYZ(w * 0.5, maxCasework.Z, d * 0.5);
                //XYZ midPoint = (minCalloutPoint + maxCalloutPoint) * 0.5;
                //XYZ secDir = XYZ.BasisY;
                //XYZ up = XYZ.BasisZ;
                //XYZ viewdir = secDir.CrossProduct(up);

                ////Transform object for bounding box
                //Transform t = Transform.Identity;
                //Transform t2 = t;

                //t.Origin = midPoint;
                //t.BasisX = secDir;
                //t.BasisY = up;
                //t.BasisZ = viewdir;
                //BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                //sectionBox.Transform = t;
                //sectionBox.Min = sectionMin;
                //sectionBox.Max = sectionMax;


                ////2nd section's box
                //XYZ secDir2 = XYZ.BasisX;
                //XYZ viewdir2 = secDir2.CrossProduct(up);
                //double w2 = maxCasework.Y - minCasework.Y;
                //double d2 = maxCasework.X - minCasework.X;
                //XYZ sectionMin2 = new XYZ(-d2 * 0.5, 0, 0);
                //XYZ sectionMax2 = new XYZ(d2 * 0.5, maxCasework.Z, w2 * 0.5);
                //t2.Origin = midPoint;
                //t2.BasisX = secDir2;
                //t2.BasisY = up;
                //t2.BasisZ = viewdir2;
                //BoundingBoxXYZ sectionBox2 = new BoundingBoxXYZ();
                //sectionBox2.Transform = t2;
                //sectionBox2.Min = sectionMin2;
                //sectionBox2.Max = sectionMax2;
                #endregion



                //Get No-Title Viewport Type
                #region RETRIVE VIEWPORT TYPE WITH NO TITLE
                FilteredElementCollector viewPortCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports);

                Element noTitleViewPortType = viewPortCollector
                    .Cast<Viewport>()
                    .Where(x => x.Name.ToLower().Contains("no title")).FirstOrDefault();
                ElementId noTitleId = noTitleViewPortType.GetTypeId();


                if (noTitleViewPortType == null)
                {
                    TaskDialog.Show("Null Error", "No viewport found with type name containing : no title");
                    goto skipTool;
                }

                //Viewport placement on Sheet
                XYZ axonPlace = new XYZ(0, 1, 0);
                XYZ planCalloutPlace = new XYZ(1, 1, 0);
                XYZ keyPlanPlace = new XYZ(2.9, 1, 0);
                XYZ sectionPlace = new XYZ(2.9, 2, 0);
                XYZ sectionPlace2 = new XYZ(2.9, 3, 0);
                #endregion




                ViewSheet newSheetHolder = null;

                //ALL TRANSACTIONS
                using (Transaction trans = new Transaction(doc, "Pack B Sheets"))
                {
                    trans.Start();

                    //create a new sheet
                    ViewSheet newSheet = ViewSheet.Create(doc, packBTitleBlockId);
                    newSheet.Name = sheetNameNumber[0];
                    newSheet.SheetNumber = sheetNumber;
                    newSheet.get_Parameter(sheetCategoryGuid).Set(wwSheetCategory);
                    newSheet.get_Parameter(sheetSubCategoryGuid).Set(wwSheetSubCategory);
                    newSheet.get_Parameter(sheetSeriesGuid).Set(wwSheetSer);
                    newSheet.get_Parameter(sheetIssuancePackBGuid).Set(wwSheetIss);

                    //Creating a 3D view
                    View3D view = View3D.CreateIsometric(doc, desViewFamilyType.Id);
                    view.SetSectionBox(offsetBoundingBox);
                    view.Name = sheetNameNumber[1] + " PackB Axon - " + sheetNameNumber[0];
                    view.Scale = 50;


                    //Creating a callout view
                    View viewCallout = ViewSection.CreateCallout(doc, layoutId, calloutTypeId, minCalloutPoint, maxCalloutPoint);
                    viewCallout.Name = sheetNameNumber[1] + " PackB Callout - " + sheetNameNumber[0];


                    //Creating a keyplan view
                    View keyPlanView = doc.GetElement(layoutView.Duplicate(ViewDuplicateOption.Duplicate)) as View;
                    keyPlanView.Name = sheetNameNumber[1] + " PackB KeyPlan - " + sheetNameNumber[0];
                    keyPlanView.ChangeTypeId(keyPlanTypeId);
                    keyPlanView.ViewTemplateId = keyPlanTemplateId;
                    DetailCurve dc1 = doc.Create.NewDetailCurve(keyPlanView, line1);
                    DetailCurve dc2 = doc.Create.NewDetailCurve(keyPlanView, line2);
                    DetailCurve dc3 = doc.Create.NewDetailCurve(keyPlanView, line3);
                    DetailCurve dc4 = doc.Create.NewDetailCurve(keyPlanView, line4);

                    if (_gstyle != null && _gstyle.IsValidObject)
                    {
                        dc1.LineStyle = _gstyle;
                        dc2.LineStyle = _gstyle;
                        dc3.LineStyle = _gstyle;
                        dc4.LineStyle = _gstyle;
                    }


                    //Creating a section view
                    View sectionView = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
                    sectionView.Name = sheetNameNumber[1] + " PackB Section 1 - " + sheetNameNumber[0];
                    View sectionView2 = ViewSection.CreateSection(doc, sectionTypeId, sectionBox2);
                    sectionView2.Name = sheetNameNumber[1] + " PackB Section 2 - " + sheetNameNumber[0];


                    //Insert views to sheet
                    Viewport axonViewport = Viewport.Create(doc, newSheet.Id, view.Id, axonPlace);
                    axonViewport.ChangeTypeId(noTitleId);

                    Viewport calloutViewport = Viewport.Create(doc, newSheet.Id, viewCallout.Id, planCalloutPlace);
                    calloutViewport.ChangeTypeId(noTitleId);
                    viewCallout.Scale = 50;

                    Viewport keyPlanViewport = Viewport.Create(doc, newSheet.Id, keyPlanView.Id, keyPlanPlace);
                    keyPlanViewport.ChangeTypeId(noTitleId);

                    Viewport sectionViewport = Viewport.Create(doc, newSheet.Id, sectionView.Id, sectionPlace);
                    sectionViewport.ChangeTypeId(noTitleId);
                    sectionView.Scale = 50;

                    Viewport sectionViewport2 = Viewport.Create(doc, newSheet.Id, sectionView2.Id, sectionPlace2);
                    sectionViewport2.ChangeTypeId(noTitleId);
                    sectionView2.Scale = 50;

                    newSheetHolder = newSheet;
                    doc.Regenerate();

                    trans.Commit();
                }

                if(newSheetHolder != null)
                {
                    uidoc.ActiveView = newSheetHolder;
                }


                // Return success result
                skipTool:
                string toolName = "Pack B Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Pack B Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);

                message = e.Message;
                return Result.Failed;
            }

        }
    }
}