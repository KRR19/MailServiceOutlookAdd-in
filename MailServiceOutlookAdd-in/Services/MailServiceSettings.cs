using MailServiceOutlookAdd_in.Models;
using MailServiceOutlookAdd_in.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Remoting.Messaging;

namespace MailServiceOutlookAdd_in
{
    public class MailServiceSettings
    {
        public static string QuestionMessageBody { get; private set; } = "Ist diese E-mail bezogen auf ein Projekt?";
        public static string QuestionMessageHeader { get; private set; } = "Speichern";
        public static string INBOX_FOLDER { get; set; }
        public static string AutoMailFlag { get; private set; } = "AutoMail";
        public static string CopyMailFlag { get; private set; } = "Copy";

        public static string RootFolder { get; private set; }

        private readonly string SettingsFileName = "MailServiceOutlookAdd-inSettings.txt";


        public MailServiceSettings()
        {
            GetRootFolder();
        }

        public List<RecipientService> GetRecipientSettings()
        {
            List<RecipientService> recipientSettings = new List<RecipientService>();
            string line;
            using (StreamReader streamReader = new StreamReader(SettingsFileName))
            {
                while ((line = streamReader.ReadLine()) != null)
                {
                    RecipientService recipientSetting = new RecipientService();
                    if (line.StartsWith(nameof(recipientSetting.Subject)))
                    {

                        recipientSetting.Subject = GetValueByKey(line, nameof(recipientSetting.Subject));
                        
                        line = streamReader.ReadLine();
                        recipientSetting.Domain = GetValueByKey(line, nameof(recipientSetting.Domain));
                        
                        line = streamReader.ReadLine();
                        recipientSetting.Body = GetValueByKey(line, nameof(recipientSetting.Body));

                        recipientSettings.Add(recipientSetting);
                    }
                }
            }
            return recipientSettings;
        }
        private string GetRootFolder()
        {
            string value = String.Empty;
            string line;
            try
            {
                using (StreamReader streamReader = new StreamReader(SettingsFileName))
                {
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        value = GetValueByKey(line, nameof(RootFolder));
                        if (value != null)
                        {
                            break;
                        }
                    }
                }
            }
            catch (System.IO.FileNotFoundException)
            {
                CreateDefaultFile();
                value = GetRootFolder();
            }

            RootFolder = value;
            return value;
        }

        private string GetValueByKey(string line, string key)
        {
            line = line.Trim();
            if (!line.StartsWith(key))
            {
                return null;
            }
            return StringService.QuotesValue(line);
        }


        private void CreateDefaultFile()
        {
            using (StreamWriter sw = new StreamWriter(SettingsFileName, false, System.Text.Encoding.Default))
            {
                string defaultRootFolder = $"{nameof(RootFolder)}: 'Projekte'";
                RecipientService recipientSettings = new RecipientService
                {
                    Subject = $"{nameof(recipientSettings.Subject)}: 'New email'",
                    Domain = $"{nameof(recipientSettings.Domain)}: 'gmail.com'",
                    Body = $"{nameof(recipientSettings.Body)}: 'Please read email'"
                };

                sw.WriteLine(defaultRootFolder);
                sw.WriteLine();
                sw.WriteLine(recipientSettings.Subject);
                sw.WriteLine(recipientSettings.Domain);
                sw.WriteLine(recipientSettings.Body);
            }
        }

    }
}
