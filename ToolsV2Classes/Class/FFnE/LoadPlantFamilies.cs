using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class LoadPlantFamilies : IExternalCommand
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

                //Element collections can be done here
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                List<string> existingImageType = new List<string>();
                FilteredElementCollector imageTypeCollector = collector.OfCategory(BuiltInCategory.OST_RasterImages).WhereElementIsElementType();

                foreach (Element item in imageTypeCollector)
                {
                    existingImageType.Add(item.Name);
                }

                string familyPath = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Plants\\";
                List<string> plantFamilies = Directory.GetFiles(familyPath, "*.*", SearchOption.TopDirectoryOnly).ToList();
                int plantFamiliesCount = plantFamilies.Count();

                string imagePath = "G:\\Shared drives\\Dev-Deliverables\\Design Technology\\Revit Content\\Families\\Plants\\Images\\";




                List<string> plantImages = Directory.GetFiles(imagePath, "*.*", SearchOption.TopDirectoryOnly).ToList();

                List<string> plantImagesFileNames = new List<string>();

                foreach (var item in plantImages)
                {
                    plantImagesFileNames.Add(Path.GetFileName(item));
                }
                int loadedImageCount = 0;
                string familyNames = "";


                using (Transaction tx = new Transaction(doc, "Load Plant Families"))
                {
                    tx.Start();
                    //Any Model Edits go here

                    foreach(string familytPath in plantFamilies)
                    {
                        Family tempPlantFamily = null;
                        FamilyLoadOption newOption = new FamilyLoadOption();
                        doc.LoadFamily(familytPath, newOption, out tempPlantFamily);
                        if (tempPlantFamily != null) 
                        {
                            familyNames += tempPlantFamily.Name + Environment.NewLine;
                        }
                    }

                    int count = 0;
                    foreach (string imagetPath in plantImages)
                    {
                        if (!existingImageType.Contains(plantImagesFileNames[count]))
                        {
                            ImageTypeOptions newImageOption = new ImageTypeOptions(imagetPath, false, ImageTypeSource.Link);
                            ImageType imageTemp = ImageType.Create(doc, newImageOption);
                            loadedImageCount++;
                        }

                        count++;
                    }



                    

                    tx.Commit();
                }
                if (loadedImageCount > 0)
                {
                    TaskDialog.Show("SUCCESS", "Plant families and images are loaded succefully in Revit Project." + Environment.NewLine + familyNames + Environment.NewLine +
"Total Plant Images Loaded = " + loadedImageCount.ToString());
                }
                else
                {
                    TaskDialog.Show("SUCCESS", "Plant families and images were loaded succefully in Revit Project." + Environment.NewLine + familyNames);
                }

            //skipTool:
                string toolName = "Load Plant Families";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);
                return Result.Succeeded;

            }
            catch (Exception e)
            {
                string toolName = "Load Plant Families";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }



        //Any Methods Go HERE
        public ElementId CreateRevitLink(Document doc, string pathName)
        {
            FilePath path = new FilePath(pathName);
            RevitLinkOptions options = new RevitLinkOptions(false);
            // Create new revit link storing absolute path to file
            LinkLoadResult result = RevitLinkType.Create(doc, path, options);
            return (result.ElementId);
        }

        
    }
}
