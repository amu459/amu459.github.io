using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using ExcelDataReader;

namespace ToolsV2Classes
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class MEPBOQExport : IExternalCommand
    {
        public string totalLabels = "Total Label Exception";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;


            try
            {
                #region Ask user to select boq excel
                ToolsV2Classes.Class.BOQExportImport.BrowseExcel form3 = 
                    new ToolsV2Classes.Class.BOQExportImport.BrowseExcel(commandData);
                form3.ShowDialog();
                string boqPath = form3.boqFilePath;
                #endregion

                #region ExcelDataReader
                FileStream stream = File.Open(boqPath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;
                //1. Reading from a OpenXml Excel file (2007+ format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                System.Data.DataTable hvacSheet = result.Tables[1];
                System.Data.DataTable h75fSheet = result.Tables[2];
                System.Data.DataTable bmsSheet = result.Tables[3];
                System.Data.DataTable eleSheet = result.Tables[4];
                System.Data.DataTable feSheet = result.Tables[6];
                System.Data.DataTable ffSheet = result.Tables[8];
                System.Data.DataTable elvSheet = result.Tables[9];
                #endregion
                int costColumn = 8;


                //HVAC Cost Pushing
                List<BuiltInCategory> hvacCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctLinings,
                    BuiltInCategory.OST_DuctInsulations
                };
                costColumn = 8;
                int hvacCount = CostToTemplate(hvacSheet, doc, hvacCats,
                    "HVAC Cost Import", costColumn);
                if (hvacCount == -1)
                {
                    goto cleanup;
                }


                //75F Cost Pushing
                List<BuiltInCategory> h75fCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_DuctAccessory,
                };
                costColumn = 7;
                int h75fCount = CostToTemplate(h75fSheet, doc, h75fCats,
                    "75F Cost Import", costColumn);


                //BMS Cost Pushing
                List<BuiltInCategory> bmsCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_SpecialityEquipment
                };
                costColumn = 7;
                int bmsCount = CostToTemplate(bmsSheet, doc, bmsCats,
                    "BMS Cost Import", costColumn);


                //ELE Cost Pushing
                List<BuiltInCategory> eleCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_LightingDevices,
                    BuiltInCategory.OST_LightingFixtures
                };
                costColumn = 7;
                int eleCount = CostToTemplate(eleSheet, doc, eleCats,
                    "ELE Cost Import", costColumn);


                //FE Cost Pushing
                List<BuiltInCategory> feCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_MechanicalEquipment
                };
                costColumn = 7;
                int feCount = CostToTemplate(feSheet, doc, feCats,
                    "FE Cost Import", costColumn);


                //FF Cost Pushing
                List<BuiltInCategory> ffCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Sprinklers,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_FlexPipeCurves
                };
                costColumn = 7;
                int ffCount = CostToTemplate(ffSheet, doc, ffCats,
                    "FF Cost Import", costColumn);

                //ELV Cost Pushing
                List<BuiltInCategory> elvCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_FireAlarmDevices,
                    BuiltInCategory.OST_CommunicationDevices,
                    BuiltInCategory.OST_LightingFixtures
                };
                costColumn = 7;
                int elvCount = CostToTemplate(elvSheet, doc, elvCats,
                    "ELV Cost Import", costColumn);


                TaskDialog.Show("Revit:", "No of HVAC types = "
                    + hvacCount.ToString()
                    + Environment.NewLine + "No of 75F types = " 
                    + h75fCount.ToString()
                    + Environment.NewLine + "No of BMS types = "
                    + bmsCount.ToString() 
                    + Environment.NewLine + "No of ELE types = "
                    + eleCount.ToString()
                    + Environment.NewLine + "No of FE types = "
                    + feCount.ToString() 
                    + Environment.NewLine + "No of FF types = "
                    + ffCount.ToString() 
                    + Environment.NewLine + "No of ELV types = "
                    + elvCount.ToString()                    );

                cleanup:
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                TaskDialog.Show("T O T A L     L A B E L S:", totalLabels);
                message = e.Message;
                return Result.Failed;
            }
        }


        private Dictionary<string, int> ReadExcel
            (System.Data.DataTable hvacSheet, int costColumn)
        {
            int lineItems = hvacSheet.Rows.Count;
            string label = "";
            string costString = "";
            double tempCost = 0;
            int cost = 0;
            bool isParsed = false;
            Dictionary<string, int> hvacLabelCost = 
                new Dictionary<string, int>();

            for (int i = 1; i < lineItems; i++)
            {
                label = "";
                cost = 1;
                costString = "";
                label = hvacSheet.Rows[i][1].ToString();
                costString = hvacSheet.Rows[i][costColumn].ToString();

                if (label == "" || label == null ||
                        costString == "" || costString == null)
                {
                    continue;
                }
                else
                {
                    isParsed = double.TryParse(costString, out tempCost);
                    cost = (int)tempCost;
                    totalLabels += label + Environment.NewLine;
                    hvacLabelCost.Add(label, cost);
                }
            }
            return hvacLabelCost;
        }

        private string PrintAllData(Dictionary<string, int> labelCost)
        {
            string outputCon = "";
            foreach (var xoxo in labelCost)
            {
                outputCon += xoxo.Key + " => " + xoxo.Value.ToString()
                    + Environment.NewLine;
            }

            TaskDialog.Show("Revit:", "No of Items = " + labelCost.Count().ToString()
                + Environment.NewLine + outputCon);

            return outputCon;
        }

        private void WriteDictToFile
            (Dictionary<string, int> labelCost, string path)
        {
            using (StreamWriter fileWriter = new StreamWriter(path))
            {
                foreach(var keyPair in labelCost)
                {
                    fileWriter.WriteLine("{0}: {1}", 
                        keyPair.Key, keyPair.Value.ToString());
                }
                fileWriter.Close();
            }
        }

        private List<BoqItem> FilterLineItems (List<Element> familyTypes)
        {
            //Initiating Line items list
            List<BoqItem> lineItems = new List<BoqItem>();

            //Filtering out family types with empty Type Mark values
            foreach (var type in familyTypes)
            {
                Autodesk.Revit.DB.Parameter pTemp = type
                    .get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);
                if (pTemp.AsString() == "" || pTemp.AsString() == null)
                {
                    continue;
                }
                else
                {
                    BoqItem tempItem = new BoqItem(type, pTemp.AsString());
                    lineItems.Add(tempItem);
                }
            }
            return lineItems;
        }

        private void CompareLabelTypeMark 
            (List<BoqItem> lineItems, Dictionary<string, int> labelCost)
        {
            //Comparing Labels from excel and Type Mark from Revit
            foreach (var item in lineItems)
            {
                if (labelCost.Any(p => p.Key == item.typeMark))
                {
                    var tempLabel = labelCost
                        .First(p => p.Key == item.typeMark);

                    item.cost = tempLabel.Value;
                    item.costAvailable = true;
                }
            }
        }
        private int CostTrans 
            (Document doc, List<BoqItem> lineItems, string transId)
        {
            int tempCount = 0;
            using (Transaction trans = new Transaction(doc, transId))
            {
                trans.Start();
                foreach (var item in lineItems)
                {
                    if (item.costAvailable)
                    {
                        Autodesk.Revit.DB.Parameter costParam =
                            item.typeName.LookupParameter("WW-Cost");
                        if (costParam != null)
                        {
                            if (!costParam.IsReadOnly)
                            {
                                costParam.Set((double)item.cost);
                                tempCount++;
                            }
                        }
                        else
                        {
                            TaskDialog.Show("Revit Shared Parameter Error",
                                "Tch Tch Tcha!"
                                + Environment.NewLine
                                + "WW-Cost shared parameter could not be found!");
                            tempCount = -1;
                            break;
                        }
                    }
                }
                trans.Commit();
            }
            return tempCount;
        }

        private int CostToTemplate
            (System.Data.DataTable sheetId, Document doc, 
            List<BuiltInCategory> defaultCats, string transId,
            int costColumn)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            Dictionary<string, int> labelCost = new Dictionary<string, int>();

            labelCost = ReadExcel(sheetId, costColumn);

            //fiter for categories
            ElementMulticategoryFilter catFilter =
                new ElementMulticategoryFilter(defaultCats);

            //Family types contained within Categories
            List<Element> familyTypes = collector.WherePasses(catFilter)
                .WhereElementIsElementType().ToElements().ToList();

            //Filtering out family types with empty Type Mark values
            List<BoqItem> lineItems = FilterLineItems(familyTypes);

            //Comparing Labels from excel and Type Mark from Revit
            CompareLabelTypeMark(lineItems, labelCost);

            //Pushing Cost Data to Revit file
            int tempCount = CostTrans(doc, lineItems, transId);

            return tempCount;
        }

    }
}
