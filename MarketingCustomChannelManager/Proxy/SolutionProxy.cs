using Microsoft.Xrm.Sdk;

namespace Fic.XTB.MarketingCustomChannelManager.Proxy
{
    public class SolutionProxy
    {
        public Entity Entity;
        public SolutionProxy(Entity entity)
        {
            Entity = entity;
        }

        public override string ToString()
        {
            if (Entity != null)
            {
                return (string)Entity["friendlyname"];
            }
            return base.ToString();
        }
    }
}
