using MailServiceOutlookAdd_in.Models;
using Microsoft.Office.Interop.Outlook;
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
        public Outlook.Folder StartDialogService()
        {
            DialogResult dialogResult = MessageBox.Show(MailServiceSettings.QuestionMessageBody, MailServiceSettings.QuestionMessageHeader, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
            {
                return null;
            }
            ProjectFolders selectFolders = new ProjectFolders(_Application);
            if (selectFolders.ShowDialog() == DialogResult.No)
            {
                return null;
            }
            FolderFinder folderFinder = new FolderFinder(_Application);
            Outlook.Folder selectedFolder = folderFinder.FindFolderByPath(selectFolders.SelectedFolder);
            return selectedFolder;
        }

        public void SendToRecipients(string to, RecipientService recipient)
        {
            MailItem mail = _Application.CreateItem(OlItemType.olMailItem);
            mail.To = recipient.FilterByDomain(to);
            mail.Subject = recipient.Subject;
            mail.Body = recipient.Body;
            mail.DeleteAfterSubmit = true;
            mail.FlagRequest = MailServiceSettings.AutoMailFlag;
            mail.Display(false);
            mail.GetInspector.Activate();
           // mail.Send();
        }
    }
}
