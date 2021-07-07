namespace Fic.XTB.MarketingCustomChannelManager.Proxy
{
    public class LocaleProxy
    {
        public string Lcid;
        public string DisplayName;

        public LocaleProxy(string lcid, string displayName)
        {
            Lcid = lcid;
            DisplayName = displayName;
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
