using System.Xml.Linq;

namespace XmlParser
{
    public class ReadXml
    {
        private bool IsAttributeName(XmlAttributeCollection attribs, string attribName, string attribValue)
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


        private string FindAttributeValue(XmlAttributeCollection attribs, string searchStr)
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


        private string FindRowColData(XmlNode row, string column)
        {
            foreach (XmlNode cell in row)
            {
                if (IsAttributeName(cell.Attributes, "ss:Index", column) && cell.HasChildNodes)
                {
                    return cell.FirstChild.InnerText;
                }
            }
            return "";
        }


        private string FindColumnContainingText(XmlNode row, string searchText)
        {
            foreach (XmlNode cell in row)
            {
                if (cell.HasChildNodes)
                {
                    if (cell.FirstChild.InnerText.Equals(searchText))
                    {
                        return FindAttributeValue(cell.Attributes, "ss:Index");
                    }
                }
            }
            return "";
        }


        private bool IsRowWithFirstChildText(XmlNode row, string searchText)
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
