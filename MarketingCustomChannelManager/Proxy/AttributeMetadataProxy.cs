using System.Linq;
using Microsoft.Xrm.Sdk.Metadata;

namespace Fic.XTB.MarketingCustomChannelManager.Proxy
{
    public class AttributeMetadataProxy
    {
        public AttributeMetadata AttributeMetadata;
        public AttributeMetadataProxy(AttributeMetadata metadata)
        {
            AttributeMetadata = metadata;
        }

        public override string ToString()
        {
            return $"{AttributeMetadata.DisplayName.LocalizedLabels.FirstOrDefault()?.Label ?? AttributeMetadata.LogicalName} ({AttributeMetadata.LogicalName})";
        }
    }
}
