namespace MailServiceOutlookAdd_in.Services
{
    public class StringService
    {

        public static string QuotesValue(string value)
        {
            int startQuote = value.IndexOf("'") + 1;
            int endQuote = value.IndexOf("'", startQuote + 1);
            int valueLength = endQuote - startQuote;
            return value.Substring(startQuote, valueLength).Trim();
        }
    }
}
