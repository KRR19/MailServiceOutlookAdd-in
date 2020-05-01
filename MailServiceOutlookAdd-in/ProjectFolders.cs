using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailServiceOutlookAdd_in
{
    public partial class ProjectFolders : Form
    {
        private readonly Outlook.Application _Application;
        public string SelectedFolder { get; private set; }
        public ProjectFolders(Outlook.Application application)
        {
            InitializeComponent();
            _Application = application;
            CreateFoldersTree();
            this.CenterToScreen();
            TopMost = true;
        }

        private void CreateFoldersTree()
        {
            string folderName = MailServiceSettings.RootFolder;

            Outlook.Folder inBox = (Outlook.Folder)_Application.ActiveExplorer().Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            try
            {
                Outlook.Folder folder = (Outlook.Folder)inBox.Folders[folderName];
                GetFoldersTree(folder, FoldersTreeView.Nodes);
            }
            catch
            {
                MessageBox.Show("There is no folder named " + folderName + ".", "Find Folder Name");
            }

        }

        private void GetFoldersTree(Outlook.Folder folder, TreeNodeCollection nodeCollection)
        {
            Outlook.Folders childFolders = folder.Folders;

            if (childFolders.Count > 0)
            {
                int Iterator = 0;
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    string folderName = childFolder.Name;
                    TreeNode dirNode = new TreeNode { Text = folderName };
                    nodeCollection.Add(dirNode);
                    GetFoldersTree(childFolder, nodeCollection[Iterator++].Nodes);
                }
            }
        }
        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
        }

        private void FoldersTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            SelectedFolder = e.Node.FullPath;
        }
    }
}
