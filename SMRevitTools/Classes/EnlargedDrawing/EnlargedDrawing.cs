
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
using SMRevitTools.Classes.EnlargedDrawing;
using Newtonsoft.Json.Linq;
using Autodesk.Revit.Creation;
using System.Windows.Shapes;
using Microsoft.Win32.SafeHandles;
//using static System.Windows.Forms.LinkLabel;
using System.Net;
using System.Windows;
//using Microsoft.Office.Interop.Excel;

namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class EnlargedDrawing : IExternalCommand
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
                caseworkList = EnlargedDrawingMethods.GetSelection(uidoc);
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
                    .Cast<Level>().ToList();
                    //.Where(x => !x.Name.ToLower().Contains("container")).ToList();

                List<string> geometricLevels = new List<string>();

                foreach(Level level in levels)
                {
                    string geometricLevel = level.LookupParameter("SM-Geometric Level").AsString();
                    if (geometricLevel == null)
                    {
                        geometricLevels.Add("null");
                        TaskDialog.Show("Revit", "Level: " + level.Name + " doesn't have a valid value in SM-Geometric Level paratmer. Please follow process.");
                        goto skipTool;
                    }
                    if(geometricLevel == "-10")
                    {
                        geometricLevel = "00";
                    }
                    geometricLevels.Add(geometricLevel);

                }

                if (geometricLevels.Count() != geometricLevels.Distinct().Count())
                {
                    TaskDialog.Show("Revit", "SM-Geometric Level paratmer is not unique to all the levels. Please recheck and try again.");
                    goto skipTool;
                }


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

                EnlargedDrawingForm inputWindow = new EnlargedDrawingForm(uidoc, levelNames);
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
                sheetNameNumber = EnlargedDrawingMethods.CheckSheetName(sheetNameNumber);
                if (sheetNameNumber[1] == "Unknown floor")
                {
                    goto skipTool;
                }
                int levelIndex = Array.IndexOf(levelNames, sheetNameNumber[1]);
                string selectedGeometricLevel = geometricLevels[levelIndex];
                if (sheetNameNumber[1] == "CONTAINER")
                {
                    sheetNameNumber[1] = "X Floor";
                    selectedGeometricLevel = "00";
                }


                #endregion



                //Combined Solid for entire casework selection
                #region GET BOUNDING BOX OF SELECTION

                //Solid unionSolid = null;
                //unionSolid = EnlargedDrawingMethods.GetUnionSolid(caseworkList);
                //if (unionSolid == null)
                //{
                //    // If union solid is null
                //    goto skipTool;
                //}


                //Get bounidng box of Union solid with 0.5 feet offset
                //BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
                //offsetBoundingBox = EnlargedDrawingMethods.GetBB(unionSolid);
                //if (offsetBoundingBox == null)
                //{
                //    // If bounding box is null
                //    goto skipTool;
                //}

                BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
                offsetBoundingBox = EnlargedDrawingMethods.GetDirectBoundingBox(caseworkList, doc);

                #endregion

                XYZ minCasework = offsetBoundingBox.Min;
                XYZ maxCasework = offsetBoundingBox.Max;

                XYZ centerPoint = new XYZ((maxCasework.X + minCasework.X)*0.5, (maxCasework.Y + minCasework.Y) * 0.5, minCasework.Z);



                //Required Sheet metadata
                #region SHEET METADATA

                string wwSheetCategory = "02-Document"; //Default and unchagned
                string wwSheetSubCategory = "30-Room and Area Detials"; //Default and unchagned
                string wwSheetSer = "04-Details"; //Changes as per Level selected
                //if (sheetNameNumber[1] != "00")
                //{
                //    wwSheetSer = selectedGeometricLevel + "-" + sheetNameNumber[1];
                //}

                //Get list of sheet parameters
                ViewSheet defaultSheet = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>().FirstOrDefault();
                Guid sheetCategoryGuid = defaultSheet.LookupParameter("SM-SheetCategory").GUID;
                Guid sheetSubCategoryGuid = defaultSheet.LookupParameter("SM-SheetSubCategory").GUID;
                Guid sheetSeriesGuid = defaultSheet.LookupParameter("SM-SheetSeries").GUID;


                //Calculate next Sheet number based on sheet list
                string sheetNumber = EnlargedDrawingMethods.GetNextSheetNumberIncreament(doc, selectedGeometricLevel);

                //Get Titleblock Family for package BId
                ElementId packBTitleBlockId = EnlargedDrawingMethods.GetPackBTitleBlockId(doc);

                #endregion


                //Retrive 3D View type
                #region RETRIVE 3D VIEW TYPE
                FilteredElementCollector viewTypeCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));

                ViewFamilyType desViewFamilyType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("sm-enlarged 3d axon")).FirstOrDefault();
                #endregion
                if (desViewFamilyType == null)
                {
                    TaskDialog.Show("Revit,", "Axon View type is null type is null");
                }


                //Retrive Callout Plan view
                //Check which layout plan to create callout in
                #region RETRIVE CALLOUT VIEW TYPE
                Level layoutLevel = null;
                foreach (Level level in levels)
                {
                    if (level.Name.Equals(sheetNameNumber[1]))
                    {
                        layoutLevel = level;
                        break;
                    }
                    if(level.Name.ToLower().Contains("container"))
                    {
                        layoutLevel = level;
                        break;
                    }
                }
                if (layoutLevel == null)
                {
                    TaskDialog.Show("Null Error", "No Interior plan match found");
                    goto skipTool;
                }
                FilteredElementCollector viewCollector2 = new FilteredElementCollector(doc).OfClass(typeof(View)).OfCategory(BuiltInCategory.OST_Views);
                View layoutView = viewCollector2.Cast<View>()
                    .Where(x => x.Name.Contains(sheetNameNumber[1])
                    && x.Name.Contains("Interior Layout"))
                    .FirstOrDefault();
                ElementId layoutId = layoutView.Id;
                if(layoutId == null || layoutId.Value == -1)
                {
                    TaskDialog.Show("Revit error", "Interior Layout View ID error");
                    goto skipTool;
                }
                ViewFamilyType calloutType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("doc-enlarged plan")).FirstOrDefault();
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
                    .Where(x => x.Name.ToLower().Contains("doc-key plan")).FirstOrDefault();
                ElementId keyPlanTypeId = keyPlanType.Id;

                View keyPlanTemplate = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.ToLower().Contains("doc-key plan")).FirstOrDefault();
                ElementId keyPlanTemplateId = keyPlanTemplate.Id;
                #endregion


                #region RETRIVE CEILING PLAN TYPE
                ViewFamilyType ceilingType = viewTypeCollector.Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("sm-doc-enlarged rcp")).FirstOrDefault();
                ElementId ceilingTypeId = ceilingType.Id;

                View ceilingLayoutView = viewCollector2.Cast<View>()
                    .Where(x => x.Name.Contains(sheetNameNumber[1])
                    && x.Name.Contains("Reflected Ceiling Plan"))
                    .FirstOrDefault();
                ElementId ceilingLayoutId = ceilingLayoutView.Id;
                #endregion




                //Red Lines for Key Plan
                #region SETUP REDLINES ON KEYPLAN
                var gstyles = (new FilteredElementCollector(doc)).OfClass(typeof(GraphicsStyle)).Cast<GraphicsStyle>().ToList();
                GraphicsStyle _gstyle = gstyles.Where(x => x.GraphicsStyleType == GraphicsStyleType.Projection).FirstOrDefault(x => x.Name.ToLower().Contains("solid_red - key plan"));

                FilteredElementCollector fillRegionTypes  = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType));
                Element redFilledRegion = fillRegionTypes.Where(x => x.Name.Contains("SM-Solid Red")).FirstOrDefault();


                XYZ keyPoint1 = minCalloutPoint;
                XYZ keyPoint2 = new XYZ(maxCalloutPoint.X, minCalloutPoint.Y, minCalloutPoint.Z);
                XYZ keyPoint3 = maxCalloutPoint;
                XYZ keyPoint4 = new XYZ(minCalloutPoint.X, maxCalloutPoint.Y, minCalloutPoint.Z);

                Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(keyPoint1, keyPoint2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(keyPoint2, keyPoint3);
                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(keyPoint3, keyPoint4);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(keyPoint4, keyPoint1);

                List<CurveLoop> profileloops = new List<CurveLoop>();
                CurveLoop profileLoop = new CurveLoop();
                profileLoop.Append(line1);
                profileLoop.Append(line2);
                profileLoop.Append(line3);
                profileLoop.Append(line4);

                profileloops.Add(profileLoop);
                #endregion



                //Create Section/Elevation
                #region RETRIVE SECTION TYPE
                //ViewFamilyType sectionType = viewTypeCollector
                //    .Cast<ViewFamilyType>()
                //    .Where(x => x.Name.ToLower().Contains("packbsection")).FirstOrDefault();
                //ElementId sectionTypeId = sectionType.Id;

                List<BoundingBoxXYZ> sectionBoxes = EnlargedDrawingMethods.GetViewSectionBB(maxCasework, minCasework, maxCalloutPoint, minCalloutPoint);
                BoundingBoxXYZ sectionBox = sectionBoxes[0];
                BoundingBoxXYZ sectionBox2 = sectionBoxes[1];


                //Elevation Code block
                ViewFamilyType viewElevationFamilyType = viewTypeCollector.Cast<ViewFamilyType>().Where(x => x.Name.ToLower().Contains("doc-enlarged room elevation")).FirstOrDefault();



                #endregion



                //Get No-Title Viewport Type
                #region RETRIVE VIEWPORT TYPE WITH NO TITLE
                FilteredElementCollector viewPortCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports);

                Element noTitleViewPortType = viewPortCollector
                    .Cast<Viewport>()
                    .Where(x => x.Name.ToLower().Contains("no title")).FirstOrDefault();
                ElementId noTitleId = noTitleViewPortType.GetTypeId();


                if (noTitleViewPortType == null || noTitleId == null)
                {
                    TaskDialog.Show("Null Error", "No viewport found with type name containing : no title");
                    goto skipTool;
                }





                FilteredElementCollector viewPortCollector2 = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports);

                Element withTitleViewPortType = viewPortCollector
                    .Cast<Viewport>()
                    .Where(x => x.Name.ToLower().Contains("title w line")).FirstOrDefault();
                ElementId titleId = withTitleViewPortType.GetTypeId();


                if (noTitleViewPortType == null || noTitleId == null)
                {
                    TaskDialog.Show("Null Error", "No viewport found with type name containing : no title");
                    goto skipTool;
                }





                //Viewport placement on Sheet
                XYZ axonPlace = new XYZ(2, 1.5, 0);
                XYZ planCalloutPlace = new XYZ(0.5, 1.5, 0);
                XYZ ceilingCalloutPlace = new XYZ(1.25, 1.5, 0);
                XYZ keyPlanPlace = new XYZ(2, 0.5, 0);
                XYZ sectionPlace = new XYZ(2.9, 2, 0);
                XYZ sectionPlace2 = new XYZ(2.9, 3, 0);
                XYZ elePlace0 = new XYZ(0.5, 0.9, 0);
                XYZ elePlace1 = new XYZ(1.25, 0.9, 0);
                XYZ elePlace2 = new XYZ(0.5, 0.35, 0);
                XYZ elePlace3 = new XYZ(1.25, 0.35, 0);
                #endregion




                ViewSheet newSheetHolder = null;

                //ALL TRANSACTIONS
                using (Transaction trans = new Transaction(doc, "SMTool-Enlarged Sheet"))
                {
                    trans.Start();

                    //create a new sheet
                    ViewSheet newSheet = ViewSheet.Create(doc, packBTitleBlockId);
                    newSheet.Name = sheetNameNumber[0];
                    newSheet.SheetNumber = sheetNumber;
                    newSheet.get_Parameter(sheetCategoryGuid).Set(wwSheetCategory);
                    newSheet.get_Parameter(sheetSubCategoryGuid).Set(wwSheetSubCategory);
                    newSheet.get_Parameter(sheetSeriesGuid).Set(wwSheetSer);

                    //Creating a 3D view
                    View3D axonview = View3D.CreateIsometric(doc, desViewFamilyType.Id);
                    axonview.SetSectionBox(offsetBoundingBox);
                    axonview.Name = sheetNameNumber[1] + " Enlarged Axon - " + sheetNameNumber[0];
                    axonview.Scale = 50;


                    //Creating a callout view

                    View viewCallout = ViewSection.CreateCallout(doc, layoutId, calloutTypeId, minCalloutPoint, maxCalloutPoint);
                    viewCallout.Name = sheetNameNumber[1] + " Enlarged Callout - " + sheetNameNumber[0];
                    viewCallout.Scale = 25;
                    viewCallout.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("PLAN");


                    //creating ceiling callout view
                    View viewCeilingCallout = ViewSection.CreateCallout(doc, ceilingLayoutId, ceilingTypeId, minCalloutPoint, maxCalloutPoint);
                    viewCeilingCallout.Name = sheetNameNumber[1] + " Enlarged RCP - " + sheetNameNumber[0];
                    viewCeilingCallout.Scale = 25;
                    viewCeilingCallout.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("RCP");

                    //Creating a keyplan view
                    View keyPlanView = doc.GetElement(layoutView.Duplicate(ViewDuplicateOption.Duplicate)) as View;
                    keyPlanView.Name = sheetNameNumber[1] + " Enlarged KeyPlan - " + sheetNameNumber[0];
                    keyPlanView.ChangeTypeId(keyPlanTypeId);
                    keyPlanView.ViewTemplateId = keyPlanTemplateId;

                    if(redFilledRegion != null)
                    {
                        FilledRegion filledRegion = FilledRegion.Create(doc, redFilledRegion.Id, keyPlanView.Id, profileloops);

                    }
                    else
                    {

                    }
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
                    //View sectionView = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
                    //sectionView.Name = sheetNameNumber[1] + " PackB Section 1 - " + sheetNameNumber[0];
                    //View sectionView2 = ViewSection.CreateSection(doc, sectionTypeId, sectionBox2);
                    //sectionView2.Name = sheetNameNumber[1] + " PackB Section 2 - " + sheetNameNumber[0];

                    //Creating Elevation Views
                    ElevationMarker elevationMarker = ElevationMarker.CreateElevationMarker(doc, viewElevationFamilyType.Id, centerPoint, 50);
                    ViewSection elevationView0 = elevationMarker.CreateElevation(doc, viewCallout.Id, 0);
                    elevationView0.Name = sheetNameNumber[1] + " Elevation A - " + sheetNameNumber[0];
                    elevationView0.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("ELEVATION");

                    // Access and modify crop box
                    //BoundingBoxXYZ cropBox = elevationView0.CropBox;
                    //XYZ min = cropBox.Min;
                    //XYZ max = cropBox.Max;

                    //XYZ newMin = new XYZ(min.X, min.Y, offsetBoundingBox.Min.Z );
                    //XYZ newMax = new XYZ(max.X, max.Y, offsetBoundingBox.Max.Z );

                    //cropBox.Min = newMin;
                    //cropBox.Max = newMax;
                    //elevationView0.CropBox = cropBox;



                    //elevationView0.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set("A");
                    //elevationView0.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER).Set("A");

                    ViewSection elevationView1 = elevationMarker.CreateElevation(doc, viewCallout.Id, 1);
                    elevationView1.Name = sheetNameNumber[1] + " Elevation B - " + sheetNameNumber[0];
                    elevationView1.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("ELEVATION");
                    //elevationView1.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER).Set("B");


                    ViewSection elevationView2 = elevationMarker.CreateElevation(doc, viewCallout.Id, 2);
                    elevationView2.Name = sheetNameNumber[1] + " Elevation C - " + sheetNameNumber[0];
                    elevationView2.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("ELEVATION");
                    //elevationView2.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER).Set("C");



                    ViewSection elevationView3 = elevationMarker.CreateElevation(doc, viewCallout.Id, 3);
                    elevationView3.Name = sheetNameNumber[1] + " Elevation D - " + sheetNameNumber[0];
                    elevationView3.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set("ELEVATION");
                    //elevationView0.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER).Set("D");



                    //Insert views to sheet


                    Viewport calloutViewport = Viewport.Create(doc, newSheet.Id, viewCallout.Id, planCalloutPlace);
                    calloutViewport.ChangeTypeId(titleId);
                    viewCallout.Scale = 25;

                    Viewport calloutCeilingViewport = Viewport.Create(doc, newSheet.Id, viewCeilingCallout.Id, ceilingCalloutPlace);
                    calloutCeilingViewport.ChangeTypeId(titleId);
                    viewCeilingCallout.Scale = 25;

                    Viewport axonViewport = Viewport.Create(doc, newSheet.Id, axonview.Id, axonPlace);
                    axonViewport.ChangeTypeId(noTitleId);

                    Viewport keyPlanViewport = Viewport.Create(doc, newSheet.Id, keyPlanView.Id, keyPlanPlace);
                    keyPlanViewport.ChangeTypeId(noTitleId);

                    //Viewport sectionViewport = Viewport.Create(doc, newSheet.Id, sectionView.Id, sectionPlace);
                    //sectionViewport.ChangeTypeId(noTitleId);
                    //sectionView.Scale = 50;

                    //Viewport sectionViewport2 = Viewport.Create(doc, newSheet.Id, sectionView2.Id, sectionPlace2);
                    //sectionViewport2.ChangeTypeId(noTitleId);
                    //sectionView2.Scale = 50;

                    Viewport eleViewport0 = Viewport.Create(doc, newSheet.Id, elevationView0.Id, elePlace0);
                    eleViewport0.ChangeTypeId(titleId);
                    
                    elevationView0.Scale = 25;

                    Viewport eleViewport1 = Viewport.Create(doc, newSheet.Id, elevationView1.Id, elePlace1);
                    eleViewport1.ChangeTypeId(titleId);
                    elevationView1.Scale = 25;

                    Viewport eleViewport2 = Viewport.Create(doc, newSheet.Id, elevationView2.Id, elePlace2);
                    eleViewport2.ChangeTypeId(titleId);
                    elevationView2.Scale = 25;

                    Viewport eleViewport3 = Viewport.Create(doc, newSheet.Id, elevationView3.Id, elePlace3);
                    eleViewport3.ChangeTypeId(titleId);
                    elevationView3.Scale = 25;



                    newSheetHolder = newSheet;
                    doc.Regenerate();

                    trans.Commit();
                }

                if (newSheetHolder != null)
                {
                    uidoc.ActiveView = newSheetHolder;
                }


            // Return success result
            skipTool:
                string toolName = "Enlarged Sheet";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Enlarged Sheet";
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