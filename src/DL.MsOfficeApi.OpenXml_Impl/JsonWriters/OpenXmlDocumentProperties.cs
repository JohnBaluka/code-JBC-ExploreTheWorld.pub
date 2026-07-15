using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Maps OPC core/extended/custom file properties to the document-property lists the
    // VBA writers emit. The factory receives (name, value, msoPropertyType) so each app
    // area can create its own entity type.
    internal static class OpenXmlDocumentProperties
    {
        // MsoDocProperties values
        private const int msoPropertyTypeNumber = 1;
        private const int msoPropertyTypeBoolean = 2;
        private const int msoPropertyTypeDate = 3;
        private const int msoPropertyTypeString = 4;
        private const int msoPropertyTypeFloat = 5;

        // app.xml element local names mapped to the VBA built-in property names.
        private static readonly (string ElementName, string PropertyName, int Type)[] ExtendedPropertyNames =
        {
            ("Template", "Template", msoPropertyTypeString),
            ("Application", "Application name", msoPropertyTypeString),
            ("Manager", "Manager", msoPropertyTypeString),
            ("Company", "Company", msoPropertyTypeString),
            ("TotalTime", "Total editing time", msoPropertyTypeNumber),
            ("Pages", "Number of pages", msoPropertyTypeNumber),
            ("Words", "Number of words", msoPropertyTypeNumber),
            ("Characters", "Number of characters", msoPropertyTypeNumber),
            ("CharactersWithSpaces", "Number of characters (with spaces)", msoPropertyTypeNumber),
            ("Lines", "Number of lines", msoPropertyTypeNumber),
            ("Paragraphs", "Number of paragraphs", msoPropertyTypeNumber),
            ("Slides", "Number of slides", msoPropertyTypeNumber),
            ("Notes", "Number of notes", msoPropertyTypeNumber),
            ("HiddenSlides", "Number of hidden Slides", msoPropertyTypeNumber),
            ("MMClips", "Number of multimedia clips", msoPropertyTypeNumber),
            ("PresentationFormat", "Presentation format", msoPropertyTypeString),
            ("Bytes", "Number of bytes", msoPropertyTypeNumber),
            ("HyperlinkBase", "Hyperlink base", msoPropertyTypeString),
        };

        public static List<T> BuildBuiltIn<T>(IPackageProperties core, ExtendedFilePropertiesPart? extended,
            Func<string, string?, int, T> factory)
        {
            static string? Date(DateTime? value) =>
                value?.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);

            var list = new List<T>
            {
                factory("Title", core.Title, msoPropertyTypeString),
                factory("Subject", core.Subject, msoPropertyTypeString),
                factory("Author", core.Creator, msoPropertyTypeString),
                factory("Keywords", core.Keywords, msoPropertyTypeString),
                factory("Comments", core.Description, msoPropertyTypeString),
                factory("Last author", core.LastModifiedBy, msoPropertyTypeString),
                factory("Revision number", core.Revision, msoPropertyTypeString),
                factory("Creation date", Date(core.Created), msoPropertyTypeDate),
                factory("Last save time", Date(core.Modified), msoPropertyTypeDate),
                factory("Last print date", Date(core.LastPrinted), msoPropertyTypeDate),
                factory("Category", core.Category, msoPropertyTypeString),
                factory("Content status", core.ContentStatus, msoPropertyTypeString),
            };

            foreach (var (name, value, type) in ReadExtendedProperties(extended))
            {
                list.Add(factory(name, value, type));
            }

            return list;
        }

        public static List<T> BuildCustom<T>(CustomFilePropertiesPart? part, Func<string, string?, int, T> factory)
        {
            var list = new List<T>();
            if (part == null) return list;

            try
            {
                using var reader = new StreamReader(part.GetStream());
                var root = XDocument.Parse(reader.ReadToEnd()).Root;
                if (root == null) return list;

                foreach (var property in root.Elements().Where(e => e.Name.LocalName == "property"))
                {
                    string name = property.Attribute("name")?.Value ?? string.Empty;
                    var valueElement = property.Elements().FirstOrDefault();

                    string? value = valueElement?.Value;
                    int type = valueElement?.Name.LocalName switch
                    {
                        "i4" or "int" or "i8" => msoPropertyTypeNumber,
                        "bool" => msoPropertyTypeBoolean,
                        "filetime" or "date" => msoPropertyTypeDate,
                        "r4" or "r8" => msoPropertyTypeFloat,
                        _ => msoPropertyTypeString,
                    };

                    list.Add(factory(name, value, type));
                }
            }
            catch
            {
                // Unreadable custom properties keep whatever was parsed so far.
            }

            return list;
        }

        // Reads the value of a single named custom property ("_MarkAsFinal" etc.).
        public static string? GetCustomPropertyValue(CustomFilePropertiesPart? part, string name)
        {
            string? value = null;

            BuildCustom(part, (propertyName, propertyValue, _) =>
            {
                if (propertyName == name) value = propertyValue;
                return propertyName;
            });

            return value;
        }

        private static List<(string Name, string? Value, int Type)> ReadExtendedProperties(ExtendedFilePropertiesPart? part)
        {
            var list = new List<(string, string?, int)>();
            if (part == null) return list;

            try
            {
                using var reader = new StreamReader(part.GetStream());
                var root = XDocument.Parse(reader.ReadToEnd()).Root;
                if (root == null) return list;

                foreach (var (elementName, propertyName, type) in ExtendedPropertyNames)
                {
                    string? value = root.Elements().FirstOrDefault(e => e.Name.LocalName == elementName)?.Value;
                    if (string.IsNullOrEmpty(value)) continue;

                    list.Add((propertyName, value, type));
                }
            }
            catch
            {
                // Unreadable extended properties keep whatever was parsed so far.
            }

            return list;
        }
    }
}
