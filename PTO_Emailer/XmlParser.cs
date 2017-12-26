using System.Xml;

namespace PTO_Emailer
{
    class XmlParser
    {
        public static bool IsAttributeName(XmlAttributeCollection attribs, string attribName, string attribValue)
        {
            foreach (XmlAttribute attrib in attribs)
            {
                if (attrib.Name.Equals(attribName) && attrib.Value.Equals(attribValue))
                {
                    return true;
                }
            }
            return false;
        }


        public static string FindAttributeValue(XmlAttributeCollection attribs, string searchStr)
        {
            foreach (XmlAttribute attrib in attribs)
            {
                if (attrib.Name.Equals(searchStr))
                {
                    return attrib.Value;
                }
            }
            return "";
        }


        public static string FindRowColData(XmlNode row, string column)
        {
            foreach (XmlNode cell in row)
            {
                if (XmlParser.IsAttributeName(cell.Attributes, "ss:Index", column) && cell.HasChildNodes)
                {
                    return cell.FirstChild.InnerText;
                }
            }
            return "";
        }


        public static string FindColumnContainingText(XmlNode row, string searchText)
        {
            foreach (XmlNode cell in row)
            {
                if (cell.HasChildNodes)
                {
                    if (cell.FirstChild.InnerText.Equals(searchText))
                    {
                        return XmlParser.FindAttributeValue(cell.Attributes, "ss:Index");
                    }
                }
            }
            return "";
        }


        public static bool IsRowWithFirstChildText(XmlNode row, string searchText)
        {
            if (row.HasChildNodes)
            {
                if (row.FirstChild.InnerText.Equals(searchText))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
