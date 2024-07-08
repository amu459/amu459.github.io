using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ToolsV2Classes
{
    public class UIOperation
    {

        //private const string AirtableApiKey = "patIRKsQT59ISawn2.aadc263d3b3d4d372cb3cd2c443c50d9dce3753fbcda76242d5f006af9081426";
        //private const string BaseId = "appKEjw9GdgNbUOzi";
        private const string AirtableApiKey = "patuk6S2DoCjt54Jk.74b085aa1478ee0f938e75218ee8897ded9ac9594a56103f4ec9f5be5582a4d8";
        private const string BaseId = "appEDwukHaDxQnQLI";
        private const string TableName = "Item list";
        // private const string FieldName = "Generic Name ";

        public void calculateBOQ(Document doc, string filePath, UIApplication uIApplication)

        {

            DateTime startTime = DateTime.Now;
            try
            {
                int notInAirtableCount = 0;
                int totalElemCount = 0;

                if (filePath != "")
                {

                    List<FamilyCostInfoClass> familyCostInfoClassList = new List<FamilyCostInfoClass>();

                    familyCostInfoClassList = UtilityClass.GetAirtableValue(AirtableApiKey, BaseId, TableName);

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    ICollection<Element> placedElementList = collector.OfClass(typeof(FamilyInstance)).ToElements();
                    placedElementList = placedElementList.Where(element => (element as FamilyInstance).Symbol.Name.Contains("WWI")).ToList();
                    List<string> PlacedFamilyNameList = new List<string>();
                    List<FamilyInstanceBOQClass> actualFamilyCostInfoClassList = new List<FamilyInstanceBOQClass>();
                    //placedElementList = placedElementList.GroupBy(x => (x as FamilyInstance).Symbol.Name).Select(g => g.FirstOrDefault()).ToList(); ;
                    foreach (Element element in placedElementList)
                    {
                        string TypeName = (element as FamilyInstance).Symbol.Name;
                        Parameter parameter = element.LookupParameter("Level");
                        Level level = (doc.GetElement(parameter.AsElementId()) as Level);
                        bool isValidLevel = true;
                        if (level == null)
                        {
                            isValidLevel = false;
                        }
                        else if (level.Name.ToString() == "CONTAINER LEVEL")
                        {
                            isValidLevel = false;
                        }
                        //else { isValidLevel = true; }

                        if (isValidLevel)
                        {
                            totalElemCount++;
                            if (!PlacedFamilyNameList.Contains(TypeName))
                            {


                                FamilyInstanceBOQClass familyBOQtInfoClass = new FamilyInstanceBOQClass();

                                if (familyCostInfoClassList.FirstOrDefault(x => x.TypeName == TypeName) == null)
                                {
                                    familyBOQtInfoClass.FamilyType = TypeName;
                                    familyBOQtInfoClass.count = 1;
                                    familyBOQtInfoClass.FamilyName = UtilityClass.getFamilyName(element);
                                    familyBOQtInfoClass.cost = 0;
                                    familyBOQtInfoClass.familyGenericName = "-";
                                    familyBOQtInfoClass.approvalStatus = "-";
                                    familyBOQtInfoClass.Category = UtilityClass.getCategory(element);
                                    familyBOQtInfoClass.SubCategory = "-";
                                    familyBOQtInfoClass.Remarks = "Not in Airtable";
                                    notInAirtableCount++;
                                    actualFamilyCostInfoClassList.Add(familyBOQtInfoClass);
                                    PlacedFamilyNameList.Add(TypeName);
                                }
                                else
                                {
                                    familyBOQtInfoClass.FamilyType = (element as FamilyInstance).Symbol.Name;
                                    try
                                    {

                                        familyBOQtInfoClass.count = 1;
                                        familyBOQtInfoClass.FamilyName = UtilityClass.getFamilyName(element);
                                        familyBOQtInfoClass.SubCategory = familyCostInfoClassList.Where(x => x.TypeName == familyBOQtInfoClass.FamilyType).OrderByDescending(x => x.cost).FirstOrDefault().SubCategory;
                                        familyBOQtInfoClass.Category = UtilityClass.getCategory(element);
                                        familyBOQtInfoClass.Remarks = familyCostInfoClassList.Where(x => x.TypeName == familyBOQtInfoClass.FamilyType).OrderByDescending(x => x.cost).FirstOrDefault().Remarks;
                                        familyBOQtInfoClass.cost = familyCostInfoClassList.Where(x => x.TypeName == familyBOQtInfoClass.FamilyType).OrderByDescending(x => x.cost).FirstOrDefault().cost;
                                        familyBOQtInfoClass.familyGenericName = familyCostInfoClassList.FirstOrDefault(x => x.TypeName == familyBOQtInfoClass.FamilyType).familyGenericName;
                                        familyBOQtInfoClass.approvalStatus = familyCostInfoClassList.FirstOrDefault(x => x.TypeName == familyBOQtInfoClass.FamilyType).approvalStatus;
                                        actualFamilyCostInfoClassList.Add(familyBOQtInfoClass);
                                        PlacedFamilyNameList.Add(TypeName);

                                    }
                                    catch (Exception ex)
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                if (familyCostInfoClassList.FirstOrDefault(x => x.TypeName == TypeName) == null)
                                {
                                    actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).count++;
                                }
                                else
                                {
                                    double cost = actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).cost;
                                    int unitCount = actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).count;
                                    double unitCost = cost / unitCount;
                                    actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).cost =
                                        actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).cost + unitCost;

                                    actualFamilyCostInfoClassList.FirstOrDefault(x => x.FamilyType == TypeName).count++;
                                }
                            }
                        }

                    }
                    //string fileName =  "\\BOQ.xlsx";
                    string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    string fileName = "\\"+doc.Title+"_BOQ_"+ timeStamp + ".xlsx";
                    UtilityClass.ListToExcel(actualFamilyCostInfoClassList, filePath, fileName);
                    TaskDialogShowClass taskDialogShowClass = new TaskDialogShowClass();

                    taskDialogShowClass = UtilityClass.readExcel(filePath, fileName);
                    taskDialogShowClass.totalNonAirtableTypeCount = notInAirtableCount.ToString();
                    taskDialogShowClass.totalElementCount = totalElemCount.ToString();
                    object[] args = new object[] { taskDialogShowClass.totalCost, taskDialogShowClass.totalElementCount, taskDialogShowClass.totalNonAirtableTypeCount };
                    TaskDialog.Show("BOQ Export", string.Format("Bulk BOQ Export Successful. \n\n Total Cost = Rs. {0} \n Total Bulk Items = {1} \n Non Standard Item Types = {2}", args));
                }
                else
                {
                    TaskDialog.Show("Error Message", "Please provide path for the BOQ Export");
                }
                string toolName = "BulkExport";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Success", doc, uIApplication, detlaMilliSec);
            }
            catch (Exception e)
            {

                string toolName = "BulkExport";
                DateTime endTime = DateTime.Now;
                var deltaTime = endTime - startTime;
                var detlaMilliSec = deltaTime.Milliseconds;
                HelperClassLibrary.logger.CreateDump(toolName, "Failure - " + e.Message, doc, uIApplication, detlaMilliSec);

                TaskDialog.Show("Error Message", e.Message);
            }

        }

        public void browsePath(UserControl1 userControl1)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                userControl1.ExcelPath.Text = folderDlg.SelectedPath;
            }
        }

    }
}
