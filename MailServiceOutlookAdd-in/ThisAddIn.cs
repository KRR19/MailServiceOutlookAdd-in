﻿using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
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
            MailServiceSettings mailServiceSettings = new MailServiceSettings();
            try
            {
                Outlook.Folder a = (Outlook.Folder)Application.ActiveExplorer().Session.DefaultStore.GetRootFolder().Folders[MailServiceSettings.RootProjectFolderName];
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

        private void OutlookApplication_ItemSend(object Item, ref bool Cancel)
        {
            StartService(Item);
        }

        private void Items_ItemAdd(object Item)
        {
            StartService(Item);
        }

        private void StartService(object item)
        {
            MailItem mailItem = (MailItem)item;
            MailService mailService = new MailService(OutlookApplication);
            mailService.StartService(mailItem);
            InternalStartup();
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
                    MailService mailService = new MailService(OutlookApplication);
                    mailService.StartService(mailItem);
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