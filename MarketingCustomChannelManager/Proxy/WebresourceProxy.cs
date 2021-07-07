using Microsoft.Xrm.Sdk;

namespace Fic.XTB.MarketingCustomChannelManager.Proxy
{
    public class WebresourceProxy
    {
        public Entity Entity;
        public WebresourceProxy(Entity entity)
        {
            Entity = entity;
        }

        public WebresourceProxy() { }

        public override string ToString()
        {
            if (Entity != null)
            {
                return (string)Entity["name"];
            }
            return base.ToString();
        }
    }
}
