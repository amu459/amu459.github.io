using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ToolsV2Classes
{
    public partial class GetLevelNo : System.Windows.Forms.Form
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;

        public string levelNoInput;
        public GetLevelNo(ExternalCommandData commandData)
        {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void LevelNoBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void continueButton_Click(object sender, EventArgs e)
        {
            levelNoInput = LevelNoBox.Text;
            continueButton.DialogResult = DialogResult.OK;
            Debug.WriteLine("OK");
            Close();
            return;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            cancelButton.DialogResult = DialogResult.Cancel;
            Debug.WriteLine("Canceled");
        }
    }
}
