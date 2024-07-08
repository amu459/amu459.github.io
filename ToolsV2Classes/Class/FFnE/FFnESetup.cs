using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows;
using ToolsV2Classes.Class.FFnE;
using ToolsV2Classes.Class.PackB;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
//using ToolsV2Classes.Class.FFnE.FFnE_Airtable;
using HelperClassLibrary.Airtable;
using Microsoft.VisualBasic;
using System.Windows.Controls;

namespace ToolsV2Classes
{

    [TransactionAttribute(TransactionMode.Manual)]
    public class FFnESetup : IExternalCommand
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

                string spaceTypeName;
                List<string> selectedLevels = new List<string>();
                List<Element> furnitureList = new List<Element>();


                //Collect elements from user selection, filter out unrequired elements
                #region Get SpaceType Info YOU GET: "string spaceTypeName" and "List<string> selectedLevels"

                //COLLECT ELEMENTS FROM SELECTION
                furnitureList = packBmethods.GetSelection(uidoc);
                if (0 == furnitureList.Count())
                {
                    // If no casework/furniture are selected.
                    goto skipTool;
                }

                //GET LIST OF LEVELS FROM REVIT
                string[] levelNames = GetLevelNames(doc);

                //Ask for Casework Name and select Floor
                FFnESpaceTypeWindow inputWindow = new FFnESpaceTypeWindow(uidoc, levelNames);
                inputWindow.ShowDialog();
                spaceTypeName = inputWindow.inputText;
                selectedLevels = inputWindow.inputLervelNamesSelected;

                List<string> SpaceObjInput = new List<string>();
                SpaceObjInput.Add(spaceTypeName);

                #endregion





                //Create Key Plan and Zoomed view to export image from Revit
                #region GET BOUNDING BOX OF SELECTION

                BoundingBoxXYZ offsetBoundingBox = packBmethods.GetDirectBoundingBox(furnitureList, doc);
                XYZ minCasework = offsetBoundingBox.Min;
                XYZ maxCasework = offsetBoundingBox.Max;

                //Get min and max point for callout views
                XYZ minCalloutPoint = new XYZ(minCasework.X - 1, minCasework.Y - 1, minCasework.Z);
                XYZ maxCalloutPoint = new XYZ(maxCasework.X + 1, maxCasework.Y + 1, minCasework.Z);
                #endregion




                #region RETRIVE CALLOUT VIEW TYPE
                FilteredElementCollector viewTypeCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));
                View layoutView = doc.ActiveView;
                ElementId layoutId = layoutView.Id;
                ViewFamilyType calloutType = viewTypeCollector
                    .Cast<ViewFamilyType>()
                    .Where(x => x.Name.ToLower().Contains("layout")).FirstOrDefault();
                ElementId calloutTypeId = calloutType.Id;
                #endregion



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


                //Standard SKU families to be added on Loose furniture table with count
                //That record id to be added to Space Type

                View viewCallout;
                View keyPlanView;




                //If Furniture List contains standard spacetypes
                string standardFamilyName = CheckIfStandardSpace(doc, furnitureList);

                if(standardFamilyName != null)
                {
                    List<Element> allStandardFamilies = GetAllStandardSpaces(doc, standardFamilyName, layoutView);

                    //Get bounding box for each
                    List<BoundingBoxXYZ> standardBoundingBoxes = GetAllBBox(allStandardFamilies, doc);

                    //Get red line outline for each
                    List<Line> redLines = GetAllRedLines(standardBoundingBoxes);



                    TaskDialog.Show("Revit win", standardFamilyName + " Is Standard Space type - Count = " + Environment.NewLine
                        + allStandardFamilies.Count().ToString());
                    using (Transaction taxFraud = new Transaction(doc, "Standard FFnE"))
                    {
                        taxFraud.Start();
                        //Creating a callout view
                        viewCallout = ViewSection.CreateCallout(doc, layoutId, calloutTypeId, minCalloutPoint, maxCalloutPoint);
                        viewCallout.Name = "FFnE Export - " + selectedLevels[0] + " - " + spaceTypeName;



                        //Creating a keyplan view
                        keyPlanView = doc.GetElement(layoutView.Duplicate(ViewDuplicateOption.Duplicate)) as View;
                        keyPlanView.Name = "FFnE KeyPlan - " + selectedLevels[0] + " - " + spaceTypeName;
                        keyPlanView.ChangeTypeId(keyPlanTypeId);
                        keyPlanView.ViewTemplateId = keyPlanTemplateId;

                        foreach(Line ln in redLines)
                        {
                            DetailCurve dc = doc.Create.NewDetailCurve(keyPlanView, ln);
                            if (_gstyle != null && _gstyle.IsValidObject)
                            {
                                dc.LineStyle = _gstyle;
                            }
                        }

                        taxFraud.Commit();
                    }
                    goto StandardCommited;

                }
                else
                {
                    goto nonStandard;
                }






                nonStandard:
                using (Transaction taxFraud = new Transaction(doc, "FFnE"))
                {
                    taxFraud.Start();
                    //Creating a callout view
                    viewCallout = ViewSection.CreateCallout(doc, layoutId, calloutTypeId, minCalloutPoint, maxCalloutPoint);
                    viewCallout.Name = "FFnE Export - " + selectedLevels[0] + " - " + spaceTypeName;



                    //Creating a keyplan view
                    keyPlanView = doc.GetElement(layoutView.Duplicate(ViewDuplicateOption.Duplicate)) as View;
                    keyPlanView.Name = "FFnE KeyPlan - " + selectedLevels[0] + " - " + spaceTypeName;
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



                    taxFraud.Commit();
                }
                StandardCommited:
                #region Misc tasks
                ElementId invalid = ElementId.InvalidElementId;
                IList<ElementId> invalidList = new List<ElementId>();
                invalidList.Add(invalid);
                uidoc.Selection.SetElementIds(invalidList);


                uidoc.ActiveView = viewCallout;
                string savedCalloutPath = ExportToImage3(doc);

                uidoc.ActiveView = keyPlanView;
                string savedKeyplanPath = ExportToImage3(doc);

                uidoc.ActiveView = layoutView;


                using (Transaction delViews = new Transaction(doc, "FFnE - Cleaning up your mess"))
                {
                    delViews.Start();
                    doc.Delete(viewCallout.Id);
                    doc.Delete(keyPlanView.Id);

                    delViews.Commit();
                }



                    //upload to imgur for getting web url of image

                    string keyPlanUrl = UploadImageToImgur(savedKeyplanPath).Result;
                string calloutPlanUrl = UploadImageToImgur(savedCalloutPath).Result;

                SpaceObjInput.Add(calloutPlanUrl);
                SpaceObjInput.Add(keyPlanUrl);

                //CREATE AIRTABLE ROOT OBJECT
                SpaceTypeObj.Root newRoot = SpaceTypeObj(SpaceObjInput, selectedLevels[0]);

                //SEND TO AIRTABLE BASE MENTIONED IN PROJECT INFORMATION
                string baseId = GetAirtbaleBaseId(doc, "WW-MasterScheduleAirtable");
                string pushResult = SendSpaceTypeData.CreateAirtableRecordSpaceType(newRoot, baseId).Result;

                TaskDialog.Show("RevitWindow2", "Upload to Airtable Was " + pushResult);

                //Task.Delay(2000);
                //DELETE LOCAL IMAGES
                DeleteImage(savedCalloutPath);
                DeleteImage(savedKeyplanPath);

                #endregion


            skipTool:
                string toolName = "FFnE Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "FFnE Setup";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }



        //Any Methods Go HERE

        static string ExportToImage3(Document doc)
        {
            string filepath;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Export Image");

                string desktop_path = Environment.GetFolderPath(
                  Environment.SpecialFolder.Desktop);

                View view = doc.ActiveView;

                filepath = Path.Combine(desktop_path,
                  view.Name);

                ImageExportOptions img = new ImageExportOptions();

                img.ZoomType = ZoomFitType.Zoom;
                img.Zoom = 100;
                img.PixelSize = 1024;
                img.ImageResolution = ImageResolution.DPI_600;
                img.FitDirection = FitDirectionType.Horizontal;
                img.ExportRange = ExportRange.CurrentView;
                img.HLRandWFViewsFileType = ImageFileType.JPEGLossless;
                img.FilePath = filepath;
                img.ShadowViewsFileType = ImageFileType.JPEGLossless;

                doc.ExportImage(img);

                tx.RollBack();

                filepath = Path.ChangeExtension(
                  filepath, "jpg");

                //Process.Start(filepath);

            }
            return filepath;
        }

        static void DeleteImage(string imagePath)
        {
            if (File.Exists(imagePath))
            {
                File.Delete(imagePath);
            }
        }


        static string[] GetLevelNames(Document doc)
        {
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

            return levelNames;
        }

        private static async Task<string> UploadImageToImgur(string imagePath)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Client-ID", "a3a9a8bd2854be1");

            var imageContent = new ByteArrayContent(File.ReadAllBytes(imagePath));
            var form = new MultipartFormDataContent
        {
            { imageContent, "image", Path.GetFileName(imagePath) }
        };

            var response = client.PostAsync("https://api.imgur.com/3/upload", form).Result;
            var responseContent = response.Content.ReadAsStringAsync().Result;

            var jsonResponse = JObject.Parse(responseContent);
            var imageUrl = jsonResponse["data"]["link"].ToString();
            //TaskDialog.Show("Rev win", imageUrl);
            Task.Delay(2000);


            return imageUrl;
        }

        public static SpaceTypeObj.Root SpaceTypeObj(List<string> SpaceObjInput, string floor)
        {
            //SpaceObjInput = spaceName, sapceImage, keyplanImage
            List<SpaceTypeObj.Record> recordsList = new List<SpaceTypeObj.Record>();

            //foreach (string floor in floorNos)
            //{
            //    SpaceTypeObj.Record record = new SpaceTypeObj.Record();
            //    record = record.CreateSpaceObjRecord(SpaceObjInput, floor);
            //    records.Add(record);
            //}

            SpaceTypeObj.Record tempRecord = new SpaceTypeObj.Record();
            recordsList.Add(tempRecord.CreateSpaceObjRecord(SpaceObjInput, floor));

            SpaceTypeObj.Root root = new SpaceTypeObj.Root();
            root.records = recordsList;
            root.Typecast = true;
            return root;
        }


        public static string GetAirtbaleBaseId(Document doc, string paramName)
        {
            string baseId = "";
            // Get the Project Information element
            ProjectInfo projectInfo = doc.ProjectInformation;

            // Check if the project information element is valid
            if (projectInfo == null)
            {
                throw new InvalidOperationException("Project Information element not found.");
            }

            // Get the parameter by name
            Parameter parameter = projectInfo.LookupParameter(paramName);

            // Check if the parameter is found
            if (parameter == null)
            {
                throw new InvalidOperationException($"Parameter '{paramName}' not found in Project Information.");
            }

            baseId = parameter.AsString();

            return baseId;
        }




        public static string CheckIfStandardSpace(Document doc, List<Element> furnitureList)
        {
            string matchedFamily = null;
            List<FamilyInstance> furnitureSystem = new List<FamilyInstance>();
            foreach(Element el in furnitureList)
            {
                if (el is FamilyInstance)
                {
                    furnitureSystem.Add(el as FamilyInstance);
                }
            }

            List<string> listOfStandardSpaces = new List<string>()
            {
                "IN-Furniture-Meet-SmallAV-4P",
                "IN-Furniture-Meet-MediumAV-6P",
                "IN-Furniture-Meet-LargeAV-10P",
                "IN-Furniture-Work-Exec",
                "IN-Phonebooth",
                "IN-Furniture-We-Printnook"
            };

            if (furnitureSystem.Count()  > 0)
            {
                foreach (FamilyInstance fi in furnitureSystem)
                {
                    string familyInstanceName = fi.Symbol.FamilyName;


                    foreach(string stdspc in listOfStandardSpaces)
                    {
                        if (familyInstanceName.Contains(stdspc))
                        {
                            matchedFamily = stdspc;
                            break;
                        }
                    }

                }
            }

            return matchedFamily;
        }


        public static List<Element> GetAllStandardSpaces (Document doc, string furnitureSystemName, View layoutView)
        {
            List<Element> furnitureSystems = new List<Element>();
            FilteredElementCollector newCollector = new FilteredElementCollector(doc, layoutView.Id);
            List<FamilyInstance> familyInstances = newCollector.OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>().Where(x => x.Symbol.FamilyName.Contains(furnitureSystemName)).ToList();

            furnitureSystems = familyInstances.Cast<Element>().ToList();

            return furnitureSystems;
        }


        public static List<BoundingBoxXYZ> GetAllBBox (List<Element> furnitureList, Document doc)
        {
            List<BoundingBoxXYZ> bBoxList= new List<BoundingBoxXYZ>();
            foreach (Element element in furnitureList)
            {
                List<Element> temList = new List<Element>();
                temList.Add(element);
                bBoxList.Add(packBmethods.GetDirectBoundingBox(temList, doc));
            }
            return bBoxList;
        }

        public static List<Line> GetAllRedLines (List<BoundingBoxXYZ> bbList) 
        {
            List<Line> redLines = new List<Line>();
            foreach (var bBox in bbList)
            {
                XYZ minCasework = bBox.Min;
                XYZ maxCasework = bBox.Max;
                //Get min and max point for callout views
                XYZ minCalloutPoint = new XYZ(minCasework.X - 1, minCasework.Y - 1, minCasework.Z);
                XYZ maxCalloutPoint = new XYZ(maxCasework.X + 1, maxCasework.Y + 1, minCasework.Z);

                XYZ keyPoint1 = minCalloutPoint;
                XYZ keyPoint2 = new XYZ(maxCalloutPoint.X, minCalloutPoint.Y, minCalloutPoint.Z);
                XYZ keyPoint3 = maxCalloutPoint;
                XYZ keyPoint4 = new XYZ(minCalloutPoint.X, maxCalloutPoint.Y, minCalloutPoint.Z);

                Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(keyPoint1, keyPoint2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(keyPoint2, keyPoint3);
                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(keyPoint3, keyPoint4);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(keyPoint4, keyPoint1);

                redLines.Add(line1);
                redLines.Add(line2);
                redLines.Add(line3);
                redLines.Add(line4);
            }

            return redLines;
        }
        //public static List<Line>


    }
}



