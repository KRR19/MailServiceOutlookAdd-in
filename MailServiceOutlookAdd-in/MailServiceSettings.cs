using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace MailServiceOutlookAdd_in
{
    public class MailServiceSettings
    {
        public static string QuestionMessageBody { get; private set; } = "Ist diese E-mail bezogen auf ein Projekt?";
        public static string QuestionMessageHeader { get; private set; } = "Speichern";
        public static string INBOX_FOLDER { get; set; }

        public static string RootProjectFolderName { get; private set; }

        public MailServiceSettings()
        {
            GetSettingsFromFile();
        }

        private string GetSettingsFromFile()
        {
            string value = String.Empty;
            try
            {
                using (StreamReader streamReader = new StreamReader("RootFolderConfige.txt"))
                {
                    value = streamReader.ReadToEnd();
                    value.Trim();
                }
            }
            catch (System.IO.FileNotFoundException)
            {
                CreateDefaultFile();
                value = GetSettingsFromFile();
            }

            RootProjectFolderName = value;
            return value;
        }

        private void CreateDefaultFile()
        {
            using (StreamWriter sw = new StreamWriter("RootFolderConfige.txt", false, System.Text.Encoding.Default))
            {
                string defaultRootFolder = "Projekte";
                sw.Write(defaultRootFolder);
            }
        }

    }
}
