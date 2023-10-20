using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using HelperClassLibrary;
using System.Windows.Controls;
using HelperClassLibrary.Airtable;
using HelperClassLibrary.Airtable.SignTypeObjectClasses;
//using HelperClassLibrary.Airtable.MainObjectClasses;
using HelperClassLibrary.Airtable.IconTypeObjectClasses;
using HelperClassLibrary.Airtable.ArrowTypeObjectClasses;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using System.Threading;

namespace ToolsV2Classes.Class.Revit_To_Airtable
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RevitToAT : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;

            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            try
            {
                HelperClassLibrary.Airtable.AirtableBaseLink formAirtable = new HelperClassLibrary.Airtable.AirtableBaseLink(commandData);
                formAirtable.ShowDialog();
                string baseLink = formAirtable.airtableBaseLink;

            
                string baseId = baseLink.Substring(21,17);

                if (baseId == "" || baseId == "failed" || baseId == null || !baseId.StartsWith("app"))
                {
                    TaskDialog.Show("🥪 Stupid Sandwich 🥪", "💀 Base ID bro!!?? 💀" + Environment.NewLine
                        + "Aborting :/");
                    goto cleanup;
                }
                //Retrive all the record IDs from other tables
                SignageRoot.Root signageTypesObj = GetAirtableData.GetSignageTypesObjAsync(baseId).Result;
                IconRoot.Root IconTypesObj = GetAirtableData.GetIconObjAsync(baseId).Result;
                ArrowRoot.Root ArrowTypesObj = GetAirtableData.GetArrowObjAsync(baseId).Result;





                //GET ALL WAYFINDING VIEWS ID
                FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
                List<Autodesk.Revit.DB.View> wayfindingViews = viewCollector
                    .OfCategory(BuiltInCategory.OST_Views)
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(x => x.Name.ToLower().Contains("wayfinding plan"))
                    .Where(x => !x.Name.ToLower().Contains("x floor")).ToList();

                if (wayfindingViews.Count() < 1)
                {
                    TaskDialog.Show("How Did This Happen??", "Wayfinding views not found in model. Reach out to VDC team. Perhaps old Revit template.");
                    goto cleanup;
                }
                List<SignageObject> sixRightSignType = new List<SignageObject>();
                List<SignageObject> sixLeftSignType = new List<SignageObject>();
                List<SignageObject> threeRightSignType = new List<SignageObject>();
                List<SignageObject> threeLeftSignType = new List<SignageObject>();
                List<SignageObject> dirSignType = new List<SignageObject>();
                List<SignageObject> bladeSignType = new List<SignageObject>();


                List<MainRoot.Record> mainRootRecordObject = new List<MainRoot.Record>();
                // Update each signage for sku
                using (Transaction transaction = new Transaction(doc, "Wayfinding Signage <> Airtable"))
                {
                    transaction.Start();
                    int countSKU = 1;
                    foreach (Autodesk.Revit.DB.View wayfView in wayfindingViews)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc, wayfView.Id);

                        List<SignageObject> tempObj = new List<SignageObject>();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Directional Signage-6 Line-Right Arrow", countSKU);
                        sixRightSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();

                        tempObj.Clear();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Directional Signage-6 Line-Left Arrow", countSKU);
                        sixLeftSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();

                        tempObj.Clear();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Directional Signage-3 Line-Right Arrow", countSKU);
                        threeRightSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();


                        tempObj.Clear();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Directional Signage-3 Line-Left Arrow", countSKU);
                        threeLeftSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();


                        tempObj.Clear();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Directory Signage", countSKU);
                        dirSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();

                        tempObj.Clear();
                        tempObj = GetSignageObj(collector, "WWI-Wayfinding-Blade Signage", countSKU);
                        bladeSignType.AddRange(tempObj);
                        countSKU += tempObj.Count();

                    }

                    List<List<SignageObject>> allSignTypes = new List<List<SignageObject>>
                {
                    sixRightSignType,
                    sixLeftSignType,
                    threeRightSignType,
                    threeLeftSignType,
                    dirSignType,
                    bladeSignType
                };
                    //Get list of all the records
                    mainRootRecordObject = CreateMainRootObj(signageTypesObj, IconTypesObj, ArrowTypesObj, allSignTypes);
                    transaction.Commit();
                }
                int sixRightSignCount = sixRightSignType.Count();
                int sixLeftSignCount = sixLeftSignType.Count();
                int threeRightSignCount = threeRightSignType.Count();
                int threeLeftSignCount = threeLeftSignType.Count();
                int dirSignCount = dirSignType.Count();
                int bladeSignCount = bladeSignType.Count();

                double totalRecords = mainRootRecordObject.Count();
                int rootObjectsRequired = (int)Math.Ceiling(totalRecords / 10);
                int eachRootRecCount = (int)Math.Ceiling(totalRecords / rootObjectsRequired);

                int recordCount = 0;
                //Divide records into list of records with maximum 10 records each
                List<List<MainRoot.Record>> mainRootObjectList = new List<List<MainRoot.Record>>(rootObjectsRequired + 1);
                for(int i = 0; i < rootObjectsRequired+1; i++)
                {
                    List<MainRoot.Record> tempRecList = new List<MainRoot.Record>();
                    for(int j = 0; j< eachRootRecCount; j++)
                    {
                        recordCount++;
                        if (recordCount > totalRecords)
                        {
                            break;
                        }
                        tempRecList.Add(mainRootRecordObject[i * (eachRootRecCount) + j]);

                    }
                    mainRootObjectList.Add(tempRecList);
                    if (recordCount > totalRecords)
                    {
                        break;
                    }
                }

                List<MainRoot.Root> rootList = new List<MainRoot.Root>();
                int tempRecCount = 0;
                foreach(List<MainRoot.Record> root in  mainRootObjectList)
                {
                    MainRoot.Root tempRoot = new MainRoot.Root();
                    List<MainRoot.Record> tempRecordList = new List<MainRoot.Record>();
                    foreach (MainRoot.Record record in root)
                    {
                        tempRecordList.Add(record);
                        tempRecCount++;
                    }
                    tempRoot.records = tempRecordList;
                    rootList.Add(tempRoot);
                }


                int recCount = 0;

                string resultDia = "";
                foreach (MainRoot.Root newRoot in rootList)
                {
                    recCount += newRoot.records.Count();
                    resultDia += "_" + SetAirtableData.CreateAirtableRecord(baseId, newRoot).Result + " _ " + Environment.NewLine;
                    Task.Delay(1000);
                }

                TaskDialog.Show("Congratulations?", "Data push to airtable status code : " + Environment.NewLine + resultDia
                    + Environment.NewLine + "Total Records Pushed = " + recCount.ToString()
                    + Environment.NewLine + "Directory Signage count = " + dirSignCount.ToString()
                    + Environment.NewLine + "Directional Signage 650 x 350 L count = " + sixLeftSignCount.ToString()
                    + Environment.NewLine + "Directional Signage 650 x 200 L count = " + threeLeftSignCount.ToString()
                    + Environment.NewLine + "Directional Signage 650 x 350 R count = " + sixRightSignCount.ToString()
                    + Environment.NewLine + "Directional Signage 650 x 200 R count = " + threeRightSignCount.ToString()
                    + Environment.NewLine + "Blade Signage count = " + bladeSignCount.ToString());

            // Return success result
            cleanup:
                string toolName = "Wayfinding Automation";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uiApp, detlaMilliSec);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                string toolName = "Wayfinding Automation";
                UIApplication uiApp = commandData.Application;
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uiApp, detlaMilliSec);
                message = e.Message;
                return Result.Failed;
            }

        }




        //to create signage objects from annotation families in Revit wayfinding views
        public List<SignageObject> GetSignageObj(FilteredElementCollector annoCollector, string familyName, int countSku)
        {
            List<FamilyInstance> signagesInstances = annoCollector
                    .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                    .Cast<FamilyInstance>()
                    .Where(x => x.Symbol.FamilyName.Contains(familyName))
                    .ToList();

            List<SignageObject> signageObjects = new List<SignageObject>();

            int count = countSku;
            foreach (FamilyInstance fi in signagesInstances)
            {
                signageObjects.Add(new SignageObject().CreateSignageObj(fi, count));
                count++;
                #region Showdialog
                //TaskDialog.Show("Revit:", "Total 6 line right = " + sixLineRightTypes.Count().ToString()
                //            + sixLineRight.Text01 + " <hspace> " + sixLineRight.CnfRoomName
                //            + Environment.NewLine + sixLineRight.Text02
                //            + Environment.NewLine + sixLineRight.Text03
                //            + Environment.NewLine + sixLineRight.Text04
                //            + Environment.NewLine + sixLineRight.Text05
                //            + Environment.NewLine + sixLineRight.Text06
                //            + Environment.NewLine + Environment.NewLine
                //            + sixLineRight.DirectionArrow); 
                #endregion
            }

            return signageObjects;
        }


        //to create main root object from the signage objects
        public List<MainRoot.Record> CreateMainRootObj(SignageRoot.Root signageTypesObj,
                IconRoot.Root IconTypesObj, ArrowRoot.Root ArrowTypesObj, List<List<SignageObject>> allSignTypes)
        {
            List<MainRoot.Record> mainRootRecordList = new List<MainRoot.Record>();
            foreach (List<SignageObject> signTypes in allSignTypes)
            {
                foreach (SignageObject signage in signTypes)
                {
                    MainRoot.Record tempRec = new MainRoot.Record();
                    mainRootRecordList.Add(tempRec.CreateSignageObjRecord(signageTypesObj, IconTypesObj, ArrowTypesObj, signage));

                }
            
            }
            return mainRootRecordList;
        }

    }

}
