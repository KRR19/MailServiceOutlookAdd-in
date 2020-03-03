using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailServiceOutlookAdd_in
{
    class MailService
    {
        private readonly Outlook.Application _Application;

        public MailService()
        {
        }

        public MailService(Outlook.Application application)
        {
            _Application = application;
        }
        public void StartService(Outlook.MailItem item)
        {
            DialogResult dialogResult = MessageBox.Show(MailServiceSettings.QuestionMessageBody, MailServiceSettings.QuestionMessageHeader, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
            {
                return;
            }
            ProjectFolders selectFolders = new ProjectFolders(_Application);
            if (selectFolders.ShowDialog() == DialogResult.No)
            {
                return;
            }
            FolderFinder folderFinder = new FolderFinder(_Application);
            Outlook.Folder selectedFolder = folderFinder.FindFolderByPath(selectFolders.SelectedFolder);
            item.Move(selectedFolder);
        }

    }
}
