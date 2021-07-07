using System;
using System.IO;
using System.Linq;
using ExCSS;

namespace Fic.XTB.MarketingCustomChannelManager.Model
{
    public class TileCss
    {
        public string TileName { get; set; }

        public string TileBackgroundColor { get; set; }
        public string IconBackgroundColor { get; set; }
        public string TileBorderColor { get; set; }
        public string LeftBorderColor { get; set; }
        public string TileHoverColor { get; set; }

        public string Icon { get; set; }
        public string IconClass { get; set; }

        public string FontFamily { get; set; }
        public string FontFamilyClass { get; set; }

        private readonly Stylesheet _stylesheet;

        public TileCss(string tileName, string cssContent)
        {
            var parser = new StylesheetParser();
            _stylesheet = parser.Parse(cssContent);

            TileName = tileName;

            Icon = GetIcon();
            IconClass = GetIconClass();

            FontFamily = GetFontFamily();
            FontFamilyClass = GetFontFamilyClass();

            TileBackgroundColor = GetTileBackgroundColor();
            IconBackgroundColor = GetIconBackgroundColor();
            TileBorderColor = GetTileBorderColor();
            LeftBorderColor = GetLeftBorderColor();
        }

        public TileCss(string tileName)
        {
            TileName = tileName;

            Icon = @"\EA8F";
            FontFamily = "CRMMDL2";

            TileBackgroundColor = "#752875";
            IconBackgroundColor = "#752875";
            TileBorderColor = "#752875";
            LeftBorderColor = "#752875";
        }

        public string GenerateFileContent()
        {
            var css = GetTemplateContent();

            css = css.Replace("TILE_NAME", TileName);
            css = css.Replace("TILE-BACKGROUND-COLOR", TileBackgroundColor);
            css = css.Replace("ICON-BACKGROUND-COLOR", IconBackgroundColor);
            css = css.Replace("TILE-BORDER-COLOR", TileBorderColor);
            css = css.Replace("LEFT-BORDER-COLOR", LeftBorderColor);

            return css;
        }

        private string GetTileBackgroundColor()
        {
            var tileBackgroundColor = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains("#libraryElementCustom_"))?.Style.BackgroundColor;
            return tileBackgroundColor;
        }

        private string GetIconBackgroundColor()
        {
            var tileBackgroundColor = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains(" span.tileImageWrapper"))?.Style.BackgroundColor;
            return tileBackgroundColor;
        }

        private string GetTileBorderColor()
        {
            var tileBorderColor = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains(".tileOutline.selected"))?.Style.BorderColor;
            return tileBorderColor;
        }

        private string GetLeftBorderColor()
        {
            var leftBorderColor = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains(".tileLeftBorder"))?.Style.BorderLeftColor;
            return leftBorderColor;
        }

        private string GetIcon()
        {
            var icon = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains("Tile::before"))?.Style.Content.Replace("\"","");
            var unicodeString = (icon ?? string.Empty).Select(t => $@"\{Convert.ToUInt16(t):X4}").ToList().FirstOrDefault();
            return unicodeString;
        }

        private string GetIconClass()
        {
            var iconClass = _stylesheet.StyleRules
                .FirstOrDefault(s => s.SelectorText.Contains("Tile::before"))?.SelectorText
                .Replace("::before","")
                .Substring(1);

            return iconClass;
        }

        private string GetFontFamily()
        {
            var fontFamily = _stylesheet.StyleRules.FirstOrDefault(s => s.SelectorText.Contains("TileSymbolFont"))?.Style.FontFamily.Replace("\"", "");
            return fontFamily;
        }

        private string GetFontFamilyClass()
        {
            var fontFamilyClass = _stylesheet.StyleRules
                .FirstOrDefault(s => s.SelectorText.Contains("TileSymbolFont"))?.SelectorText
                .Substring(1);

            return fontFamilyClass;
        }

        private string GetTemplateContent()
        {
            var stream = this.GetType().Assembly.GetManifestResourceStream("Fic.XTB.MarketingCustomChannelManager.Templates.template.css");
            var reader = new StreamReader(stream);
            var css = reader.ReadToEnd();

            return css;
        }
    }
}
