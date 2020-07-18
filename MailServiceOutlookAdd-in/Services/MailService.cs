
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailServiceOutlookAdd_in
{
    class MailService
    {
        private readonly Outlook.Application _Application;

        public MailService(Outlook.Application application)
        {
            _Application = application;
        }

        public void SendMail(MailItem mailItem)
        {
            if (mailItem.FlagRequest == MailServiceSettings.CopyMailFlag)
            {
                MAPIFolder folder = _Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                mailItem.Move(folder);
                return;
            }
            if (mailItem.FlagRequest != MailServiceSettings.AutoMailFlag)
            {

                string to = $"{mailItem.To}; {mailItem.CC}";

                Outlook.Folder selectedFolder = StartDialogService();
                if (selectedFolder != null)
                {
                    mailItem.SaveSentMessageFolder = selectedFolder;
                    mailItem.Save();
                    MailItem copyMail = mailItem.Copy();
                    copyMail.FlagRequest = MailServiceSettings.CopyMailFlag;
                    copyMail.Save();
                    SendToRecipients(to);
                }
            }
            mailItem.Send();
        }

        public string IncomeMail(MailItem mailItem)
        {
            mailItem.UnRead = true;
            mailItem.Save();
            
            Outlook.Folder selectedFolder = StartDialogService();
            if (selectedFolder != null)
            {
                MailItem copyMail = mailItem.Copy();
                copyMail.Move(selectedFolder);
            }
            return mailItem.EntryID;
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

        public void SendToRecipients(string to)
        {

            MailItem mail = _Application.CreateItem(OlItemType.olMailItem);
            mail.To = to;
            mail.Subject = MailServiceSettings.Subject;
            mail.Body = MailServiceSettings.Body;
            mail.DeleteAfterSubmit = true;
            mail.FlagRequest = MailServiceSettings.AutoMailFlag;

            mail.Send();
        }
    }
}
