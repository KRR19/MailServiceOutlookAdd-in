using MailServiceOutlookAdd_in.Services;

namespace MailServiceOutlookAdd_in.Models
{
    public class RecipientService
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Domain { get; set; }

        public string FilterByDomain(string to)
        {
            string filteredAddresses = string.Empty;
            while (!string.IsNullOrWhiteSpace(to))
            {
                string adress = StringService.QuotesValue(to);
                if(adress.EndsWith(Domain))
                {
                    filteredAddresses += "'" + adress + "'; ";
                }
                int delimiterIdx = to.IndexOf(';');
                if(delimiterIdx == -1)
                {
                    break;
                }
                to = to.Substring(0, delimiterIdx);
            }
            return filteredAddresses;
        }
    }
}
