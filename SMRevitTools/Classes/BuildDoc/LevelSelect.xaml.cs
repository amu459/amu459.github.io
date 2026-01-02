using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;

namespace SMRevitTools.Classes.BuildDoc
{
    /// <summary>
    /// Interaction logic for LevelSelect.xaml
    /// </summary>
    public partial class LevelSelect : Window
    {
        public UIDocument uidoc { get; set;  }
        public Document doc { get; set; }
        public string inputText { get; set; }

        public List<string> inputLevelNamesSelected { get; set; }

        public List<string> floorList;
        public List<BD_Floors> floorObjList;

        public List<string> outputFloorList;






        public LevelSelect(UIDocument uiDoc, string[] levelNames)
        {
            uidoc = uiDoc;
            doc = uiDoc.Document;
            InitializeComponent();

            floorObjList = new List<BD_Floors>();
            floorList = new List<string>();
            AddElementsInList(levelNames);
            LevelListbox.ItemsSource = floorObjList;

            //BindFloorDropDown();

        }


        private void AddElementsInList(string[] levelNames)
        {

            List<string> levelNamesList = levelNames.ToList();
            int i = 0;
            foreach (string levelName in levelNamesList)
            {
                floorList.Add(levelName);
                BD_Floors bD_Floors = new BD_Floors();
                bD_Floors.Floor_Name = levelName;
                bD_Floors.Floor_ID = i;
                floorObjList.Add(bD_Floors);
                i++;
            }
        }


        private void CancelTask(object sender, RoutedEventArgs e)
        {
            inputText = "cancel";
            Close();
            return;
        }

        //private void BuildDocumentClick(object sender, RoutedEventArgs e)
        //{
        //    List<string> testLevelList = new List<string>();
        //    foreach (var floor in floorList)
        //    {
        //        if (floor.Check_Status == true)
        //        {
        //            testLevelList.Add(floor.Floor_Name);
        //        }
        //    }
        //    inputLevelNamesSelected = testLevelList;
        //    Close();
        //    return;
        //}


        private void LevelListbox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void BuildDocumentClick(object sender, RoutedEventArgs e)
        {

            //int i = 0;
            //foreach(ListBoxItem item in LevelListbox.Items )
            //{
            //    if(item.IsSelected == true)
            //    {
            //        outputFloorList.Add(item.ToString());
            //    }
            //}
            //while (i < LevelListbox.Items.Count)
            //{
            //    // Get item's ListBoxItem
            //    ListBoxItem lbi = (ListBoxItem)LevelListbox.ItemContainerGenerator.ContainerFromIndex(i);
            //    lbi.IsSelected = true;
            //    i += 2;
            //}
            Close();
            return;
        }
    }

    public class BD_Floors
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
