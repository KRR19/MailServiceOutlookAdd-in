using Microsoft.Office.Interop.Outlook;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailServiceOutlookAdd_in
{
    public partial class ThisAddIn
    {
        public Outlook.Application OutlookApplication;
        public Outlook.Inspectors OutlookInspectors;
        private Outlook.Items Items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {

                OutlookApplication = Application as Outlook.Application;
                OutlookInspectors = OutlookApplication.Inspectors;
                OutlookInspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(OpenNewMailItem);

                Outlook.Folder sentItems = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                Items = sentItems.Items;
                Items.ItemAdd += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
            }
            catch (System.Exception exception)
            {
                MessageBox.Show($"{exception.Message} \n {exception.StackTrace}", "Achtung!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Items_ItemAdd(object Item)
        {            
            MailItem mailItem = (MailItem)Item;

            string mails = mailItem.To + mailItem.CC;
            if (MailServiceSettings.ActiveForEmail.Any(activeMail => mails.Contains(activeMail)))
            {
                MailService mailService = new MailService(OutlookApplication);
                mailService.SendMail(mailItem);
                return;
            }
            mailItem.Send();
        }


        private void OpenNewMailItem(Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem == null)
            {
                return;
            }
            Outlook.MAPIFolder parentFolder = mailItem.Parent as Outlook.MAPIFolder;
            string mailItemFolderName = parentFolder.Name;
            MailServiceSettings.INBOX_FOLDER = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Name;
            if (mailItemFolderName == MailServiceSettings.INBOX_FOLDER)
            {
                MailService mailService = new MailService(OutlookApplication);
                mailService.IncomeMail(mailItem);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
