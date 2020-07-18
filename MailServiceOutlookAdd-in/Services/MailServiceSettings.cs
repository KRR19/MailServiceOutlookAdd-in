using Newtonsoft.Json;
using System.IO;

namespace MailServiceOutlookAdd_in
{
    public class MailServiceSettings
    {
        [JsonProperty]
        public static string RootFolder { get; set; } = "Projekte";
        [JsonProperty]
        public static string Subject { get; set; } = "New email";
        [JsonProperty]
        public static string Body { get; set; } = "Please read email";
        [JsonProperty]
        public static string[] ActiveForEmail = { "mail@mail.com", "mail2@mail.com" };
        public static string QuestionMessageBody { get; private set; } = "Ist diese E-mail bezogen auf ein Projekt?";
        public static string QuestionMessageHeader { get; private set; } = "Speichern";
        public static string INBOX_FOLDER { get; set; }
        public static string AutoMailFlag { get; private set; } = "AutoMail";
        public static string CopyMailFlag { get; private set; } = "Copy";


        private Microsoft.Office.Interop.Outlook.Application _Application;

        public static readonly string SettingsFileName = @"C:\MailServiceOutlookAdd-inSettings\MailServiceOutlookAdd-inSettings.json";
        public MailServiceSettings()
        {
        }

        public MailServiceSettings(Microsoft.Office.Interop.Outlook.Application application)
        {
            _Application = application;
        }


        static MailServiceSettings()
        {
            GetSettingsFromFile();
        }

        private static void GetSettingsFromFile()
        {
            if (!File.Exists(SettingsFileName))
            {
                CreateJson();
            }

            string jsonFileString = File.ReadAllText(SettingsFileName);
            MailServiceSettings mailServiceSettings = JsonConvert.DeserializeObject<MailServiceSettings>(jsonFileString);
        }

        private static void CreateJson()
        {
            MailServiceSettings mailServiceSettings = new MailServiceSettings();
            string jsonString = JsonConvert.SerializeObject(mailServiceSettings, Formatting.Indented);
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFileName));
            using (StreamWriter sw = new StreamWriter(SettingsFileName))
            {
                sw.Write(jsonString);
                sw.Close();
            }
        }

    }
}
