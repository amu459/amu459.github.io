using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System;


namespace ToolsV2Classes.Class.FFnE
{
    /// <summary>
    /// Interaction logic for packBInputWindow.xaml
    /// </summary>
    public partial class FFnESpaceTypeWindow : Window
    {
        public UIDocument uidoc { get; }
        public Document doc { get; set; }
        public string inputText { get; set; }

        public List<string> inputLervelNamesSelected { get; set; }

        List<DDL_Floors> objFloorList;

        public FFnESpaceTypeWindow(UIDocument uiDoc, string[] levelNames)
        {
            uidoc = uiDoc;
            doc = uiDoc.Document;
            InitializeComponent();

            objFloorList = new List<DDL_Floors>();
            AddElementsInList(levelNames);
            BindFloorDropDown();
            //exCombo.ItemsSource = levelNames;
            Title = "Input Casework Name";
        }



        private void CancelTask(object sender, RoutedEventArgs e)
        {
            inputText = inputCaseworkText.Text;
            inputText = "cancel";
            Close();
            return;
        }

        private void inputCaseworkText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void AddElementsInList(string[] levelNames)
        {

            List<string> levelNamesList = levelNames.ToList();
            int idCount = 1;
            foreach (string levelName in levelNamesList)
            {
                DDL_Floors obj = new DDL_Floors();
                obj.Floor_ID = idCount;
                idCount++;
                obj.Floor_Name = levelName;
                objFloorList.Add(obj);
            }
        }

        private void BindFloorDropDown()
        {
            ddlCountry.ItemsSource = objFloorList;
        }

        private void ddlCountry_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void ddlCountry_TextChanged(object sender, TextChangedEventArgs e)
        {
            ddlCountry.ItemsSource = objFloorList.Where(x => x.Floor_Name.StartsWith(ddlCountry.Text.Trim()));
        }

        //private void AllCheckbocx_CheckedAndUnchecked(object sender, RoutedEventArgs e)
        //{
        //    BindListBOX();
        //}

        //private void BindListBOX()
        //{
        //    //testListbox.Items.Clear();
        //    //foreach (var floor in objFloorList)
        //    //{
        //    //    if (floor.Check_Status == true)
        //    //    {
        //    //        testListbox.Items.Add(floor.Floor_Name);
        //    //    }
        //    //}
        //}



        private void GetPlantName(object sender, RoutedEventArgs e)
        {
            inputText = inputCaseworkText.Text;

            List<string> testLevelList = new List<string>();
            foreach (var floor in objFloorList)
            {
                if (floor.Check_Status == true)
                {
                    testLevelList.Add(floor.Floor_Name);
                }
            }
            inputLervelNamesSelected = testLevelList;
            Close();
            return;
        }
    }

    public class DDL_Floors
    {
        public int Floor_ID
        {
            get;
            set;
        }
        public string Floor_Name
        {
            get;
            set;
        }
        public Boolean Check_Status
        {
            get;
            set;
        }
    }
}
