using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolsV2Classes.Class.FFnE;
using ToolsV2Classes.Class.PackB;

namespace ToolsV2Classes
{

    [TransactionAttribute(TransactionMode.Manual)]
    public class PlantPackageSetup : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;

            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;


            try
            {
                //Collect elements from user selection, filter out unrequired elements
                #region COLLECT ELEMENTS FROM SELECTION

                List<Element> planterList = new List<Element>();
                planterList = GetSelectedPlants(uidoc);
                if (0 == planterList.Count())
                {
                    // If no casework/furniture are selected.
                    goto skipTool;
                }
                int selectedItemsCount = planterList.Count();

                #endregion


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
                #region USER INPUT FOR PLANT AREA NAME AND LEVEL

                PlantInputWindow inputWindow = new PlantInputWindow(uidoc, levelNames);
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

                BoundingBoxXYZ offsetBoundingBox = new BoundingBoxXYZ();
                offsetBoundingBox = packBmethods.GetDirectBoundingBox(planterList, doc);

                XYZ minPlanting = new XYZ(offsetBoundingBox.Min.X - 2.5, offsetBoundingBox.Min.Y - 2.5, offsetBoundingBox.Min.Z);
                XYZ maxPlanting = new XYZ(offsetBoundingBox.Max.X + 2.5, offsetBoundingBox.Max.Y + 2.5, offsetBoundingBox.Max.Z + 2.5);
                #endregion


                //Required Sheet metadata
                #region SHEET METADATA

                string wwSheetCategory = "Interiors"; //Default and unchagned
                string wwSheetSubCategory = "06 Move In Package"; //Default and unchagned
                string wwSheetSer = "0005 Plant Package"; //Changes as per Level selected
                string wwSheetIss = "•";

                //Get list of sheet parameters
                ViewSheet defaultSheet = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>().FirstOrDefault();
                Guid sheetCategoryGuid = defaultSheet.LookupParameter("WW-SheetCategory").GUID;
                Guid sheetSubCategoryGuid = defaultSheet.LookupParameter("WW-SheetSubCategory").GUID;
                Guid sheetSeriesGuid = defaultSheet.LookupParameter("WW-SheetSeries").GUID;
                Guid sheetIssuancePackBGuid = defaultSheet.LookupParameter("WW-Sheet Issuance (Package B)").GUID;

                //Calculate next Sheet number based on sheet list
                string sheetNumber = GetNextSheetNumber(doc);

                //Get Titleblock Family for package BId
                ElementId packBTitleBlockId = packBmethods.GetPackBTitleBlockId(doc);

                #endregion



                FilteredElementCollector viewTypeCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));

                //Retrive Callout Plan view
                //Check which layout plan to create callout in
                #region RETRIVE CALLOUT VIEW TYPE
                Level layoutLevel = null;
                foreach (Level level in levels)
                {
                    if (level.Name.Equals(sheetNameNumber[1]))
                    {
                        layoutLevel = level;
                    }
                }
                if (layoutLevel == null)
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
                    .Where(x => x.Name.ToLower().Contains("des-plant")).FirstOrDefault();
                ElementId calloutTypeId = calloutType.Id;
                #endregion


                XYZ minCalloutPoint = new XYZ(minPlanting.X, minPlanting.Y, minPlanting.Z);
                XYZ maxCalloutPoint = new XYZ(maxPlanting.X, maxPlanting.Y, minPlanting.Z);

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

                //Viewport placement on Sheet
                //XYZ axonPlace = new XYZ(0, 1, 0);
                XYZ planCalloutPlace = new XYZ(2.4, 1, 0);
                XYZ keyPlanPlace = new XYZ(2.9, 0.835, 0);
                XYZ schedulePlace = new XYZ(1.38, 1.856, 0);
                XYZ legendPlace = new XYZ(2.370, 0.815, 0);
                #endregion




                #region RETRIVE PLANT BOQ SCHEDULE AND LEGEND
                FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
                ViewSchedule plantSchedule = scheduleCollector
                .Cast<ViewSchedule>()
                .Where(x => x.Name.ToLower().Contains("plant boq")).FirstOrDefault();

                FilteredElementCollector legendCollector = new FilteredElementCollector(doc).OfClass(typeof(View));
                View plantLegend = legendCollector
                .Cast<View>()
                .Where(x => x.Name.ToLower().Contains("des-planting")).FirstOrDefault();
                #endregion







                ViewSheet newSheetHolder = null;

                using (Transaction tx = new Transaction(doc, "Plant Package Sheet"))
                {
                    tx.Start();

                    //create a new sheet
                    ViewSheet newSheet = ViewSheet.Create(doc, packBTitleBlockId);
                    newSheet.Name = sheetNameNumber[0];
                    newSheet.SheetNumber = sheetNumber;
                    newSheet.get_Parameter(sheetCategoryGuid).Set(wwSheetCategory);
                    newSheet.get_Parameter(sheetSubCategoryGuid).Set(wwSheetSubCategory);
                    newSheet.get_Parameter(sheetSeriesGuid).Set(wwSheetSer);
                    newSheet.get_Parameter(sheetIssuancePackBGuid).Set(wwSheetIss);


                    //Creating a callout view
                    View viewCallout = ViewSection.CreateCallout(doc, layoutId, calloutTypeId, minCalloutPoint, maxCalloutPoint);
                    viewCallout.Name = sheetNameNumber[1] + " Planting Callout - " + sheetNameNumber[0];


                    //Creating a keyplan view
                    View keyPlanView = doc.GetElement(layoutView.Duplicate(ViewDuplicateOption.Duplicate)) as View;
                    keyPlanView.Name = sheetNameNumber[1] + " Planting KeyPlan - " + sheetNameNumber[0];
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


                    //Insert views to sheet

                    Viewport calloutViewport = Viewport.Create(doc, newSheet.Id, viewCallout.Id, planCalloutPlace);
                    calloutViewport.ChangeTypeId(noTitleId);
                    viewCallout.Scale = 50;

                    Viewport keyPlanViewport = Viewport.Create(doc, newSheet.Id, keyPlanView.Id, keyPlanPlace);
                    keyPlanViewport.ChangeTypeId(noTitleId);


                    //Inset Schedule to Sheet
                    //WW-QTO-Plant BOQ
                    if(plantSchedule != null)
                    {
                        ScheduleSheetInstance.Create(doc, newSheet.Id, plantSchedule.Id, schedulePlace);

                    }




                    //Insert Legend to Sheet
                    //DES-PLANTING
                    Viewport plantLegendViewport = Viewport.Create(doc, newSheet.Id, plantLegend.Id, legendPlace);
                    plantLegendViewport.ChangeTypeId(noTitleId);

                    newSheetHolder = newSheet;
                    doc.Regenerate();

                    tx.Commit();
                }
                if (newSheetHolder != null)
                {
                    uidoc.ActiveView = newSheetHolder;
                }




            skipTool:
                string toolName = "Plant Package Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Plant Package Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }



        //Any Methods Go HERE

        public static List<Element> GetSelectedPlants (UIDocument uidoc)
        {
            //filterout Plants from user selection and return a List of Elements
            List<Element> planterList = new List<Element>();
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
                        planterList.Add(elem);
                    }
                    else if (elem.Category.Name.ToLower().Contains("planting"))
                    {
                        planterList.Add(elem);
                    }
                }
                if (0 == planterList.Count())
                {
                    // If no plants are selected.
                    TaskDialog.Show("Revit", "OOPS!" + Environment.NewLine + "Your selection doesn't contain any Plants!");
                }
            }
            return planterList;
        }

        public static string GetNextSheetNumber(Document doc)
        {
            //List of all sheets
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>().Where(x => x.SheetNumber.ToLower().Contains("w59"))
                .ToList();

            //Get Next Sheet Number in List
            int sheetIncreament = 81;
            string sheetNumber = "W59" + sheetIncreament.ToString() + ".00";
            bool sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            while (sheetNumberExists)
            {
                sheetIncreament++;
                sheetNumber = "W59" + sheetIncreament.ToString() + ".00";
                sheetNumberExists = sheets.Any(p => p.SheetNumber == sheetNumber);
            }
            return sheetNumber;
        }






    }
}

