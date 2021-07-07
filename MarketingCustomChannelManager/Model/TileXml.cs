using System.Xml.Serialization;
using System.Collections.Generic;
namespace Fic.XTB.MarketingCustomChannelManager.Model
{
    [XmlRoot(ElementName = "Definition")]
    public class Definition
    {
        [XmlAttribute(AttributeName = "icon")]
        public string Icon { get; set; }
        [XmlAttribute(AttributeName = "fontFamily")]
        public string FontFamily { get; set; }
        [XmlAttribute(AttributeName = "cssFileName")]
        public string CssFileName { get; set; }
    }

    [XmlRoot(ElementName = "ChannelProperties")]
    public class ChannelProperties
    {
        [XmlElement(ElementName = "EntityType")]
        public string EntityType { get; set; }
        [XmlElement(ElementName = "EntitySetName")]
        public string EntitySetName { get; set; }
        [XmlElement(ElementName = "TitleFieldName")]
        public string TitleFieldName { get; set; }
        [XmlElement(ElementName = "ComplianceField")]
        public string ComplianceField { get; set; }
        [XmlElement(ElementName = "LookupViewId")]
        public string LookupViewId { get; set; }
        [XmlElement(ElementName = "QuickViewFormId")]
        public string QuickViewFormId { get; set; }
    }

    [XmlRoot(ElementName = "Label")]
    public class Label
    {
        [XmlAttribute(AttributeName = "locId")]
        public string LocId { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Labels")]
    public class Labels
    {
        [XmlElement(ElementName = "Label")]
        public List<Label> Label { get; set; }
    }

    [XmlRoot(ElementName = "ResponseType")]
    public class ResponseType
    {
        [XmlElement(ElementName = "Labels")]
        public Labels Labels { get; set; }
        [XmlAttribute(AttributeName = "id")]
        public string Id { get; set; }
        [XmlAttribute(AttributeName = "custom")]
        public string Custom { get; set; }

        public ResponseType(){}

        public ResponseType(string id, List<int> lcids)
        {
            Id = id;
            Labels = new Labels {Label = new List<Label>()};

            foreach (var lcid in lcids)
            {
                Labels.Label.Add(new Label
                {
                    LocId = lcid.ToString(),
                    Text = char.ToUpper(id[0]) + id.Substring(1)
                });
            }
        }
    }

    [XmlRoot(ElementName = "ResponseTypes")]
    public class ResponseTypes
    {
        [XmlElement(ElementName = "ResponseType")]
        public List<ResponseType> ResponseType { get; set; }
    }

    [XmlRoot(ElementName = "Tooltip")]
    public class Tooltip
    {
        [XmlAttribute(AttributeName = "locId")]
        public string LocId { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Tooltips")]
    public class Tooltips
    {
        [XmlElement(ElementName = "Tooltip")]
        public List<Tooltip> Tooltip { get; set; }
    }

    [XmlRoot(ElementName = "LibraryTile")]
    public class TileXml
    {
        [XmlElement(ElementName = "Definition")]
        public Definition Definition { get; set; }
        [XmlElement(ElementName = "ChannelProperties")]
        public ChannelProperties ChannelProperties { get; set; }
        [XmlElement(ElementName = "ResponseTypes")]
        public ResponseTypes ResponseTypes { get; set; }
        [XmlElement(ElementName = "Labels")]
        public Labels Labels { get; set; }
        [XmlElement(ElementName = "Tooltips")]
        public Tooltips Tooltips { get; set; }

        public TileXml()
        {
            Definition = new Definition();
            ChannelProperties = new ChannelProperties();
            ResponseTypes = new ResponseTypes();
            Labels = new Labels();
            Tooltips = new Tooltips();
        }
    }

}