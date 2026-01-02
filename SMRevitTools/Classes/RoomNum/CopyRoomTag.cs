using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using SMRevitTools.Classes.BuildDoc;

namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CopyRoomTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Define target level name - customize as needed

                //GET LIST OF LEVELS FROM REVIT
                string[] levelNames = GetLevelNames(doc);
                //Set up Form
                LevelSelect levelSelectWindow = new LevelSelect(uidoc, levelNames);
                levelSelectWindow.ShowDialog();

                List<BD_Floors> floorObjextList = levelSelectWindow.floorObjList;
                List<string> selectedLevelNames = new List<string>();
                foreach (var item in floorObjextList)
                {
                    if (item.Check_Status == true)
                    {
                        selectedLevelNames.Add(item.Floor_Name);
                    }
                }

                if (selectedLevelNames.Count() < 1)
                {
                    TaskDialog.Show("Revit", "No Levels Selected!");
                    goto skipped;
                }

                foreach (string targetLevelName in selectedLevelNames)
                {
                    // Collect all non-template views
                    var views = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate);

                    // Find the Interior Layout view on the target level
                    var sourceView = views.FirstOrDefault(v =>
                        v.Name.ToLower().Contains("interior layout")
                        && v.GenLevel != null
                        && v.GenLevel.Name.Equals(targetLevelName, StringComparison.OrdinalIgnoreCase));

                    if (sourceView == null)
                    {
                        TaskDialog.Show("Revit", $"No Interior Layout view found for level '{targetLevelName}'.");
                        return Result.Failed;
                    }

                    var targetLevelId = sourceView.GenLevel.Id;

                    // Find other views on the same level except the source view
                    var targetViews = views.Where(v =>
                        v.GenLevel != null
                        && v.GenLevel.Id == targetLevelId
                        && v.Id != sourceView.Id).ToList();

                    // Get all RoomTag element Ids in the source view
                    var roomTagIds = new FilteredElementCollector(doc, sourceView.Id)
                        .OfCategory(BuiltInCategory.OST_RoomTags)
                        .OfClass(typeof(SpatialElementTag))
                        .Select(tag => tag.Id)
                        .ToList();

                    if (!roomTagIds.Any())
                    {
                        message = "No room tags found in source view.";
                        return Result.Failed;
                    }

                    using (Transaction transaction = new Transaction(doc, "Copy Room Tags Across Views"))
                    {
                        transaction.Start();

                        foreach (var targetView in targetViews)
                        {
                            try
                            {
                                // Copy room tags from sourceView to targetView, no additional transform
                                ElementTransformUtils.CopyElements(
                                    sourceView,
                                    roomTagIds,
                                    targetView,
                                    Transform.Identity,
                                    new CopyPasteOptions()
                                );
                            }
                            catch (Exception ex)
                            {
                                // Log or handle any exceptions per target view
                                // For now, skipping views where copying fails
                                continue;
                            }
                        }

                        transaction.Commit();
                    }
                }

            skipped:
                string toolName = "CopyRoomTagsViaCopyPaste";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var deltaMilliSec = deltaTime.Milliseconds;
                UIApplication uiApp = commandData.Application;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, deltaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "CopyRoomTagsViaCopyPaste";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var deltaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, deltaMilliSec);
                message = e.Message;
                return Result.Failed;
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
    }
}
