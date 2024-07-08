using System.Windows.Forms;

namespace BOQStandardCheck
{
    public static class StandardBOQCheckUiOperation
    {
        public static string browsePath()
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();
            string textBoxText = "";
            if (result == DialogResult.OK)
            {
                textBoxText = folderDlg.SelectedPath;
            }
            return textBoxText;
        }
        public static string browseFile()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Browse File";
            fdlg.RestoreDirectory = true;
            string textBoxText = "";
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBoxText = fdlg.FileName;
            }
            return textBoxText;
        }
    }
}
