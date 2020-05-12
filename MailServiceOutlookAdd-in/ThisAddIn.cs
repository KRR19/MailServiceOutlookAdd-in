﻿using MailServiceOutlookAdd_in.Models;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
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

        public List<RecipientService> RecipientServices;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            MailServiceSettings mailServiceSettings = new MailServiceSettings();
            RecipientServices = new List<RecipientService>();
            RecipientServices.Add(mailServiceSettings.GetRecipientSettings().FirstOrDefault());
            try
            {
                Outlook.Folder a = (Outlook.Folder)Application.ActiveExplorer().Session.DefaultStore.GetRootFolder().Folders[MailServiceSettings.RootFolder];
            }
            catch (System.Exception exception)
            {
                MessageBox.Show("Der in den Einstellungen angegebene Ordner wurde nicht gefunden!", "Achtung!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            OutlookApplication = Application as Outlook.Application;
            OutlookInspectors = OutlookApplication.Inspectors;
            OutlookInspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(OpenNewMailItem);

            Outlook.Folder sentItems = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
            Items = sentItems.Items;
            Items.ItemAdd += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);

            MailServiceSettings.INBOX_FOLDER = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Name;
        }

        private void Items_ItemAdd(object Item)
        {
            MailItem mailItem = (MailItem)Item;

            if(mailItem.FlagRequest == MailServiceSettings.CopyMailFlag)
            {
                MAPIFolder folder = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                mailItem.Move(folder);
                return;
            }
            if (mailItem.FlagRequest != MailServiceSettings.AutoMailFlag)
            {
                MailService mailService = new MailService(OutlookApplication);
                string to = mailItem.To;
                RecipientService recipient = RecipientServices.FirstOrDefault(r => r.Subject == mailItem.Subject);

                Outlook.Folder selectedFolder = mailService.StartDialogService();
                if (selectedFolder != null)
                {
                    mailItem.SaveSentMessageFolder = selectedFolder;
                    mailItem.Save();
                    MailItem copyMail = mailItem.Copy();
                    copyMail.FlagRequest = MailServiceSettings.CopyMailFlag;
                    copyMail.Save();
                }
                if (recipient != null)
                {
                    mailService.SendToRecipients(to, recipient);
                }
            }
            mailItem.Send();
        }


        private void OpenNewMailItem(Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                Outlook.MAPIFolder parentFolder = mailItem.Parent as Outlook.MAPIFolder;
                string mailItemFolderName = parentFolder.Name;

                if (mailItemFolderName == MailServiceSettings.INBOX_FOLDER)
                {
                    mailItem.UnRead = true;
                    mailItem.Save();
                    MailService mailService = new MailService(OutlookApplication);
                    Outlook.Folder selectedFolder = mailService.StartDialogService();
                    if (selectedFolder != null)
                    {
                        MailItem copyMail = mailItem.Copy();
                        copyMail.Move(selectedFolder);
                    }
                }
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
