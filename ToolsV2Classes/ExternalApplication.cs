using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace ToolsV2Classes
{
    class ExternalApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {

            string path = Assembly.GetExecutingAssembly().Location;

            string tabName = "India Tools";
            string panelName = "पटाखा";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { }

            #region MEP Views and Sheets
            RibbonPanel panelMEP = application.CreateRibbonPanel("India Tools", "For VDC");
            PushButtonData button1 = new PushButtonData("Button1", "believe", path, "ToolsV2Classes.Generate");
            //Generate views and sheets
            PushButton pushButton1 = panelMEP.AddItem(button1) as PushButton;
            pushButton1.LargeImage = FetchPngIcon("ToolsV2Classes.logos.page.png");

            #endregion

            List<RibbonPanel> panelList = application.GetRibbonPanels(tabName);
            RibbonPanel panelLL = null;
            foreach (RibbonPanel rp in panelList)
            {
                if (rp.Name == panelName)
                {
                    panelLL = rp;
                }
            }
            if (panelLL == null)
            {
                panelLL = application.CreateRibbonPanel(tabName, panelName);
            }



            #region Lighting tool buttons
            PushButtonData vData = new PushButtonData("vButtonData", "V-2.4", path, "ToolsV2Classes.VerticalLights");
            PushButtonData hData = new PushButtonData("hButtonData", "H-2.4", path, "ToolsV2Classes.HorizontalLights");

            PushButtonData vData1200 = new PushButtonData("vButtonData1200", "V-1.2", path, "ToolsV2Classes.VerticalLights1200");
            PushButtonData hData1200 = new PushButtonData("hButtonData1200", "H-1.2", path, "ToolsV2Classes.HorizontalLights1200");

            PushButtonData loungeLightButtonData = new PushButtonData("LoungeLightsData", "you CAN", path, "ToolsV2Classes.LoungeLights");


            #region Lighting Button info:
            //FOR 1200 Lights Panel
            vData1200.LargeImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png");
            vData1200.Image = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png");
            vData1200.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights.png");
            vData1200.LongDescription = "Vertical Layout with 1200mm Lights"
                + Environment.NewLine + Environment.NewLine 
                + "Model Lighting fixtures for " +
                "Program Type = 'Work' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Suspended Light-03'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Work'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";
            //panel1200Light.AddItem(vData1200);


            hData1200.LargeImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png");
            hData1200.Image = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png");
            hData1200.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights.png");
            hData1200.LongDescription = "Horizontal Layout with 1200mm Lights"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'Work' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Suspended Light-03'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Work'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";
            //panel1200Light.AddItem(hData1200);


            //FOR 2400 Lights Panel
            vData.LargeImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png");
            vData.Image = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png");
            vData.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights.png");
            vData.LongDescription = "Vertical Layout with 2400mm Lights"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'Work' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Suspended Light-03'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Work'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";
            //panel2400Light.AddItem(vData);


            hData.LargeImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png");
            hData.Image = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png");
            hData.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights.png");
            hData.LongDescription = "Horizontal Layout with 2400mm Lights"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'Work' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Suspended Light-03'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Work'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";
            //panel2400Light.AddItem(hData);


            //Lounge Lights Button
            loungeLightButtonData.LargeImage = FetchPngIcon("ToolsV2Classes.logos.light.png");
            loungeLightButtonData.Image = FetchPngIcon("ToolsV2Classes.logos.canLight56.png");
            loungeLightButtonData.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.canLightsToolTip.png");
            loungeLightButtonData.LongDescription = "Layout with Cylindrical Lights for Lounge and Corridor"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'We' or 'Circulate'" +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Suspended Light-02'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'We'/'Circulate'." + Environment.NewLine +
                "2. Click on tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.2";


            //L00 Hotizontal
            PushButtonData hL00 = new PushButtonData("hButtonDataL00", "H-L00", path, "ToolsV2Classes.L00_H")
            {
                LargeImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png"),
                Image = FetchPngIcon("ToolsV2Classes.logos.horizontalLights56.png"),
                ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.horizontalLights.png"),
                LongDescription = "Horizontal Layout with Surface mounted Linear Lights"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'Operate' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Surface Mounted Light-07'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Operate'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.2"
            };

            PushButtonData vL00 = new PushButtonData("vButtonDataL00", "V-L00", path, "ToolsV2Classes.L00_V")
            {
                LargeImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png"),
                Image = FetchPngIcon("ToolsV2Classes.logos.verticalLights56.png"),
                ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.verticalLights.png"),
                LongDescription = "Vertical Layout with Surface mounted Linear Lights"
                + Environment.NewLine + Environment.NewLine
                + "Model Lighting fixtures for " +
                "Program Type = 'Operate' " +
                "based on standard spacing logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard lighting fixture families in project " + Environment.NewLine +
                "- 'IN-Surface Mounted Light-07'"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType set to 'Operate'." + Environment.NewLine +
                "2. Click on desired tool." + Environment.NewLine +
                "3. Happy Diwali <__><__><__>" + Environment.NewLine +
                "                               |||      |||     |||"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.2"
            };

            //Light Dimension button
            PushButtonData dimensionLightButtonData = new PushButtonData("DimensionLightsData", "Dim", path, "ToolsV2Classes.DimensionLights")
            {
                //LargeImage = FetchPngIcon("ToolsV2Classes.logos.dimLights56.png"),
                Image = FetchPngIcon("ToolsV2Classes.logos.dimLights56.png"),
                ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.dimLights.png"),
                LongDescription = "Give Dimensions to Lights within Room(s)"
                + Environment.NewLine + Environment.NewLine
                + "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with Lights modelled" + Environment.NewLine +
                "2. Click on Dimension tool." + Environment.NewLine +
                "3. Profit."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.2"
            };
            #endregion


            panelLL.AddItem(loungeLightButtonData);

            List<RibbonItem> itemCanDim = new List<RibbonItem>();
            List<RibbonItem> item2400Light = new List<RibbonItem>();
            List<RibbonItem> item1200Light = new List<RibbonItem>();
            List<RibbonItem> itemL00Light = new List<RibbonItem>();
            
            

            panelLL.AddSeparator();
            item2400Light.AddRange(panelLL.AddStackedItems(vData, hData));
            item1200Light.AddRange(panelLL.AddStackedItems(vData1200, hData1200));
            
            //panelLL.AddItem(dimensionLightButtonData);
            panelLL.AddSeparator();
            itemL00Light.AddRange(panelLL.AddStackedItems( vL00, hL00, dimensionLightButtonData));
            #endregion


            #region Power and Data
            //POWER AND DATA
            RibbonPanel panelPowerData = application.CreateRibbonPanel(tabName, "Data-Pow");
            PushButtonData powerAndDataButtonData = new PushButtonData("Power & Data", "Grommet", path, "ToolsV2Classes.PowerAndData")
            {
                LargeImage = FetchPngIcon("ToolsV2Classes.logos.powerNdata.png"),
                ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.powerNdataBig.png"),
                LongDescription = "Place Power and Data Grommets under Physical Desks"
                + Environment.NewLine + Environment.NewLine
                + "Model Grommets fixtures for " +
                "Program Type = 'Work' " +
                "based on standard power and data logic."
                + Environment.NewLine + Environment.NewLine
                + "Prerequisite: " + Environment.NewLine +
                "1. Project needs to be on project template v2022.1 or higher" + Environment.NewLine +
                "2. Load Standard Grommet family in project " + Environment.NewLine +
                "- 'WWI-PowerAndData-Grommet'" + Environment.NewLine +
                "3. Physical Desk Family Name is '1_Person-Office-Desk'" + Environment.NewLine +
                "4. No Overlapping Desks and No Overlapping floors" + Environment.NewLine +
                "5. Desks are located inside of Room and are not intersecting walls/partitions" + Environment.NewLine +
                "6. Desks must be aligned properly. More than 0.5 degrees of error in desk placement will give incorrect grommet layout"
                + Environment.NewLine + Environment.NewLine +
                "Steps to Follow:" + Environment.NewLine +
                "1. Select Room(s) with WW-ProgramType  set to 'Work'." + Environment.NewLine +
                "2. Click on tool." + Environment.NewLine +
                "3. Check whether you want to delete existing Grommets from the selected Room(s)." + Environment.NewLine +
                "4. Onomatopoeia : zzzzt"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.2"
            };

            panelPowerData.AddItem(powerAndDataButtonData);
            #endregion


            #region MEP Cost
            //PushButtonData button2 = new PushButtonData("Button2", "MEP Cost Test", path, "ToolsV2Classes.MEPBOQExport");
            ////Generate views and sheets
            //PushButton pushButton2 = panelMEP.AddItem(button2) as PushButton;
            //pushButton2.LargeImage = FetchPngIcon("ToolsV2Classes.logos.page.png");
            #endregion


            #region 3D Rooms
            RibbonPanel panelProgramming = application.CreateRibbonPanel("India Tools", "3D rOOms");
            PushButtonData button3 = new PushButtonData("Button3", "Create", path, "ToolsV2Classes.RoomMassCommand");
            //Generate views and sheets
            //PushButton pushButton3 = panelProgramming.AddItem(button3) as PushButton;
            button3.LargeImage = FetchPngIcon("ToolsV2Classes.logos.createRoom56.png");
            button3.Image = FetchPngIcon("ToolsV2Classes.logos.createRoom56.png");
            button3.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.createRoomLarge.png");
            button3.LongDescription = "Creates 3D masses based on Rooms" +
                Environment.NewLine + Environment.NewLine + "Steps:" + Environment.NewLine +
                "1. Go to 3D view, Turnoff all categories except Mass" + Environment.NewLine +
                "2. Click on Create." + Environment.NewLine +
                "3. Check for any popup for failed rooms. " +
                "Create Mass for those rooms Manually otherwise you'll be spoiled by automations." + Environment.NewLine +
                "4. Offer sacrifices to Skynet."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";

            PushButtonData button4 = new PushButtonData("Button4", "Destroy", path, "ToolsV2Classes.RoomMassDelete");
            //Generate views and sheets
            //PushButton pushButton4 = panelProgramming.AddItem(button4) as PushButton;
            button4.LargeImage = FetchPngIcon("ToolsV2Classes.logos.deleteRoom56.png");
            button4.Image = FetchPngIcon("ToolsV2Classes.logos.deleteRoom56.png");
            button4.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.deleteRoomLarge.png");
            button4.LongDescription = "Deletes created 3D masses" +
                Environment.NewLine + Environment.NewLine + "Steps: " + Environment.NewLine +
                "1. Go to 3D view in which the Mass were created." + Environment.NewLine +
                "2. Click on Destroy." + Environment.NewLine +
                "3. Make sure to Manually delete masses which were Manually created."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.3";

            List<RibbonItem> item3dRooms = new List<RibbonItem>();

            item3dRooms.AddRange(panelProgramming.AddStackedItems(button3, button4));
            #endregion


            #region C&I QTO
            //QTO
            RibbonPanel panelQTO = application.CreateRibbonPanel("India Tools", "मुद्रा (beta)");
            PushButtonData button2 = new PushButtonData("Button2", "C&I qto", path, "ToolsV2Classes.ConceptQTO");
            button2.LongDescription = "Exports C&I concept quantities to standard concept budget Excel sheet" +
                Environment.NewLine + Environment.NewLine + "Steps:" + Environment.NewLine +
                "1. Close the Excel template if opened." + Environment.NewLine +
                "2. Click on C&I qto button and select the excel template by browsing." + Environment.NewLine +
                "3. Check for Export successful or failed popup. " + Environment.NewLine +
                "4. Cross check the final quantities with VDC team once as the tool is still in Beta." + Environment.NewLine +
                "5. Profit??"
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2022.4 (Beta)";
            PushButton pushButton2 = panelQTO.AddItem(button2) as PushButton;
            pushButton2.LargeImage = FetchPngIcon("ToolsV2Classes.logos.money32.png");
            pushButton2.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.leo300.png");
            #endregion


            #region Wayfinding Automation
            //Wayfinding
            RibbonPanel airtablePanel = application.CreateRibbonPanel("India Tools", "Revit <> Airtable");
            PushButtonData button6 = new PushButtonData("Button6", "AWW", path, "ToolsV2Classes.Class.Revit_To_Airtable.RevitToAT");
            button6.LongDescription = "Automated Wayfinding Workflow" +
                        Environment.NewLine + Environment.NewLine +
                        "Prerequisite: " + Environment.NewLine +
                        "1. Project needs to be on project template v2023.2 or higher" + Environment.NewLine +
                        "2. Airtable Base needs to be duplicated from standard wayfinding base and all the records from BOQ table to be deleted." + Environment.NewLine +
                        "3. Check for the release notes for detailed dos and don'ts."
                        + Environment.NewLine + Environment.NewLine +
                        "Steps:" + Environment.NewLine +
                        "1. Click on the tool."
                        + Environment.NewLine +
                        "2. Copy > Paste the airtable base weblink and submit."
                        + Environment.NewLine +
                        "3. Enjoy the blessings from Automation gods."
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, WeWork India" + Environment.NewLine +
                        "Tool Version: 2023.2";
            PushButton pushButton6 = airtablePanel.AddItem(button6) as PushButton;
            pushButton6.LargeImage = FetchPngIcon("ToolsV2Classes.logos.wayfindingLogo100.png");

            pushButton6.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.wayfindingLogo.png");
            #endregion



            #region Bulk Furniture Export
            //Bulk Furniture Export

            PushButtonData button7 = new PushButtonData("Button7", "Bulk", path, "ToolsV2Classes.SyncAirTableQuickClass");
            button7.LongDescription = "Bulk Furniture BOQ Export" +
                        Environment.NewLine + Environment.NewLine +
                        "Prerequisite: " + Environment.NewLine +
                        "1. Project needs to be on project template v2023.2 or higher" + Environment.NewLine +
                        "2. Check for the release notes for detailed dos and don'ts."
                        + Environment.NewLine + Environment.NewLine +
                        "Steps:" + Environment.NewLine +
                        "1. Click on the tool."
                        + Environment.NewLine +
                        "2. Browse/ Copy Paste the desired path for Excel Export."
                        + Environment.NewLine +
                        "3. Click on “Calculate BOQ”."
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, WeWork India" + Environment.NewLine +
                        "Tool Version: 2023.2";
            PushButton pushButton7 = airtablePanel.AddItem(button7) as PushButton;
            pushButton7.LargeImage = FetchPngIcon("ToolsV2Classes.logos.airtableBulkFurniture100.png");

            pushButton7.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.airtableBulkFurniture.png");
            #endregion



            #region Room separation cleanup
            //Room Separation
            RibbonPanel cleanup = application.CreateRibbonPanel("India Tools", "Tide");
            PushButtonData button5 = new PushButtonData("Button5", "Clean RS", path, "ToolsV2Classes.RoomSeparation");
            button5.LongDescription = "Delete unwanted room separation lines" +
                Environment.NewLine + Environment.NewLine + "Steps:" + Environment.NewLine +
                "1. Click on the tool." 
                + Environment.NewLine +
                "2. The number of lines deleted will be displayed."
                + Environment.NewLine + Environment.NewLine +
                "Compatible with Revit 2019 to 2022" + Environment.NewLine +
                "Author: VDC, WeWork India" + Environment.NewLine +
                "Tool Version: 2023.1";
            //PushButton pushButton5 = cleanup.AddItem(button5) as PushButton;
            button5.LargeImage = FetchPngIcon("ToolsV2Classes.logos.clean32.png");
            button5.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.clean.png");
            #endregion


            #region Ghosted Desk Workset Correction Tool
            //Ghosted Desk Correction Tool
            PushButtonData button8 = new PushButtonData("Button8", "GhostBusters", path, "ToolsV2Classes.GhostedDeskWorkset");
            button8.LongDescription = "Ghosted Desk Workset Correction Tool" +
                        Environment.NewLine + Environment.NewLine +
                        "Prerequisite: " + Environment.NewLine +
                        "1. Make sure model is relinquished by other users."
                        + Environment.NewLine + Environment.NewLine +
                        "Steps:" + Environment.NewLine +
                        "1. Click on the tool."
                        + Environment.NewLine +
                        "2. Boo!"
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, WeWork India" + Environment.NewLine +
                        "Tool Version: 2023.2";
            //PushButton pushButton8 = cleanup.AddItem(button8) as PushButton;
            button8.LargeImage = FetchPngIcon("ToolsV2Classes.logos.ghostBusters32.png");

            button8.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.ghostBusters100.png");
            #endregion

            SplitButtonData sb1 = new SplitButtonData("splitButton1", "split");
            SplitButton sb = cleanup.AddItem(sb1) as SplitButton;
            sb.AddPushButton(button5);
            sb.AddPushButton(button8);


            #region Pack B setup Tool
            //Package B setup tool
            RibbonPanel packBsetupPanel = application.CreateRibbonPanel("India Tools", "For ID <3");

            PushButtonData button9 = new PushButtonData("Button9", "Pack Bae", path, "ToolsV2Classes.packBSetup");
            button9.LongDescription = "Setup Package B) sheet" +
                        Environment.NewLine + Environment.NewLine +
                        "Prerequisite: " + Environment.NewLine +
                        "1. Make sure the Revit model is using project template 2023.3 or later. Reach out to VDC team if not sure."
                        + Environment.NewLine + Environment.NewLine +
                        "Steps:" + Environment.NewLine +
                        "1. Select the casework/furniture requred in the PackB sheet."
                        + Environment.NewLine +
                        "2. Click on the tool."
                        + Environment.NewLine +
                        "3. Input sheet name and select floor."
                        + Environment.NewLine +
                        "4. Rearrage the views in the sheet as required."
                        + Environment.NewLine +
                        "5. Reconfing and add or delete views as required."
                        + Environment.NewLine +
                        "6. Pack up Bae."
                        + Environment.NewLine + Environment.NewLine +
                        "Author: VDC, WeWork India" + Environment.NewLine +
                        "Tool Version: 2023.3";
            PushButton pushButton9 = packBsetupPanel.AddItem(button9) as PushButton;
            pushButton9.LargeImage = FetchPngIcon("ToolsV2Classes.logos.packB32.png");
            pushButton9.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.packBOG250.png");

            //pushButton9.ToolTipImage = FetchPngIcon("ToolsV2Classes.logos.ghostBusters100.png");
            #endregion




            return Result.Succeeded;
        }


        private System.Windows.Media.ImageSource FetchPngIcon(string embeddedPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
