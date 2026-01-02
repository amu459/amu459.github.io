using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using SMRevitTools.Classes.TagTools;


namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class DoorTagTool : IExternalCommand
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
                FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
                FilteredElementCollector activeViewCollector = new FilteredElementCollector(doc, doc.ActiveView.Id);


                
                
                
                
                //List of All Doors
                List<FamilyInstance> doorList = doorCollector.WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Doors).Cast<FamilyInstance>().Where(x => !doc.GetElement(x.LevelId).Name.ToLower().Contains("container")).ToList();

                //Door tag for ACD or FRD
                ElementType doorTagACDFRD = collector.WhereElementIsElementType().OfCategory(BuiltInCategory.OST_DoorTags).Cast<ElementType>().Where(x => x.Name.ToLower().Contains("acd/frd")).FirstOrDefault();

                //Door tag for Coding
                ElementType doorTagCoding = collector.WhereElementIsElementType().OfCategory(BuiltInCategory.OST_DoorTags).Cast<ElementType>().Where(x => x.Name.ToLower().Contains("door coding")).FirstOrDefault();

                ElementType doorTagType = doorTagCoding;

                //levels from Model - Required or not???
                //var levels = collector.OfCategory(BuiltInCategory.OST_Levels).Cast<Level>().Where(x => !x.Name.ToLower().Contains("container")).ToList();


                List<ElementId> doorNotYetTaggedElementId = new List<ElementId>();

                View activeView = doc.ActiveView;
                if (activeView.Name.ToLower().Contains("interior"))
                {
                    doorTagType = doorTagACDFRD;
                    //Door with ACD or FRD filled in
                    List<FamilyInstance> doorWithACDFRD = TagMethods.GetDoorWithACDFRD(doorList);
                    List<ElementId> doorWithACDFRDElementId = new List<ElementId>();

                    foreach (FamilyInstance door in doorWithACDFRD)
                    {
                        doorWithACDFRDElementId.Add(door.Id);
                    }


                    //List of ACD FRD tags present that are already present in the model
                    List<IndependentTag> tagList = activeViewCollector.WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DoorTags).Cast<IndependentTag>().Where(x => x.Name.ToLower().Contains("acd/frd")).ToList();
                    List<ElementId> doorAlreadyTaggedElementId = new List<ElementId>();
                    foreach (IndependentTag tag in tagList)
                    {
                        doorAlreadyTaggedElementId.Add(tag.GetTaggedLocalElements().FirstOrDefault().Id);
                    }

                    doorNotYetTaggedElementId = doorWithACDFRDElementId.Where(a => !doorAlreadyTaggedElementId.Any(b => b.Value == a.Value)).ToList();
                }
                else if (activeView.Name.ToLower().Contains("door coding"))
                {
                    //For Partition and Door Coding Plan

                    doorTagType = doorTagCoding;
                    List<ElementId> doorElementId = new List<ElementId>();
                    foreach (FamilyInstance door in doorList) 
                    {
                        doorElementId.Add(door.Id);
                    }

                    List<IndependentTag> tagList = activeViewCollector.WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DoorTags).Cast<IndependentTag>().Where(x => x.Name.ToLower().Contains("door coding")).ToList();
                    List<ElementId> doorAlreadyTaggedElementId = new List<ElementId>();
                    foreach (IndependentTag tag in tagList)
                    {
                        doorAlreadyTaggedElementId.Add(tag.GetTaggedLocalElements().FirstOrDefault().Id);
                    }

                    doorNotYetTaggedElementId = doorElementId.Where(a => !doorAlreadyTaggedElementId.Any(b => b.Value == a.Value)).ToList();

                }

                else
                {
                    TaskDialog.Show("Revit Error", "Active view is not Interior or Door Coding view. Please use this tool by opening an appropriate view (Floor Interior Plan or Floor Door Coding Plan)");
                    goto skipped;
                }


                List<FamilyInstance> doorNotYetTaggedInstance = new List<FamilyInstance>();

                foreach (ElementId doorElementId in doorNotYetTaggedElementId)
                {
                    doorNotYetTaggedInstance.Add((FamilyInstance)doc.GetElement(doorElementId));
                }




                // Create Door Tags in Interior and Door Coding
                using (Transaction transaction = new Transaction(doc, "SM-Auto Door Tag"))
                {
                    transaction.Start();
                    //Tag Interior Layout Doors
                    foreach (ElementId doorId in doorNotYetTaggedElementId)
                    {

                        IndependentTag tag = IndependentTag.Create(doc, doorTagType.Id, doc.ActiveView.Id, new Reference(doc.GetElement(doorId)), false, TagOrientation.Horizontal, (doc.GetElement(doorId).Location as LocationPoint).Point);
                    }

                    transaction.Commit();
                }

                // Return success result

                string toolName = "Tag_Door";
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
                string toolName = "Tag_Door";
                UIApplication uiApp = commandData.Application;
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
