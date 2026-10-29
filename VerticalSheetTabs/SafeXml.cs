using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace VerticalSheetTabs
{
    // Reads XML that comes from a workbook (the custom XML parts, the theme in the file), which anyone can write.
    // DTDs are refused (no entity expansion, no external files) and the size is limited.
    public static class SafeXml
    {
        // the custom XML parts (the folders, the history)
        public const int PartMaxChars = 4000000;

        // the theme in the file
        public const int ThemeMaxChars = 2000000;

        private static XmlReaderSettings Settings(int maxChars)
        {
            return new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxChars
            };
        }

        // Throws when the XML is wrong, has a DTD or is too long.
        public static XDocument Parse(string xml, int maxChars)
        {
            using (var text = new StringReader(xml))
            using (XmlReader reader = XmlReader.Create(text, Settings(maxChars)))
            {
                return XDocument.Load(reader);
            }
        }

        public static XDocument Load(Stream stream, int maxChars)
        {
            using (XmlReader reader = XmlReader.Create(stream, Settings(maxChars)))
            {
                return XDocument.Load(reader);
            }
        }
    }
}
