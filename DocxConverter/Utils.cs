using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace DocxConverter
{
    public static class XmlExtension
    {
        /// <summary>
        /// Преобразует XmlNode в XElement
        /// </summary>
        public static XElement GetXElement(this XmlNode node)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
            {
                try
                {
                    node.WriteTo(xmlWriter);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.ReadLine();
                    throw;
                }
            }
            return xDoc.Root;
        }

        /// <summary>
        /// Преобразует XElement в XmlNode 
        /// </summary>
        public static XmlNode GetXmlNode(this XElement element)
        {
            using (XmlReader xmlReader = element.CreateReader())
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlReader);
                return xmlDoc;
            }
        }

        /// <summary>
        /// Возвращает корневой узел
        /// </summary>
        public static XElement GetRootParent(this XElement element)
        {
            var root = element;
            while (root.Parent != null)
            {
                root = root.Parent;
            }
            return root;
        }

        /// <summary>
        /// Возвращает элемент-предок с заданным именем, атрибутом и значением атрибута 
        /// </summary>
        /// <param name="element">текущий элемент</param>
        /// <param name="elementName">имя элемента-предка</param>
        /// <param name="attrName">имя атрибута элемента-предка</param>
        /// <param name="attrValue">значение атрибута элемента-предка</param>
        public static XElement GetAncestor(this XElement element, string elementName, string attrName, string attrValue)
        {
            var ancestors = element.Ancestors(elementName);
            foreach (var item in ancestors)
            {
                if (item.Name == elementName)
                {
                    var attr = item.Attribute(attrName);
                    if (attr != null && attr.Value == attrValue)
                    {
                        return item;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Возвращает значение, указывающее, есть ли атрибут с заданным именеме в этом элементе.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="attributeName">Имя атрибута.</param>
        /// <returns>true если атрибут с таким именем есть; иначе false</returns>
        public static bool HasAttribute(this XElement element, XName attributeName)
        {
            return element.Attribute(attributeName) != null;
        }

        /// <summary>
        /// Возвращает значение, указывающее, есть ли в этом элементе атрибут с заданным именем и, если есть, равно ли его значение заданному.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="attributeName">Имя атрибута.</param>
        /// <param name="attributeValue">Значение атрибута.</param>
        /// <returns>true если атрибут с таким именем есть и его значение равно заданному; иначе false</returns>
        public static bool HasAttributeValue(this XElement element, XName attributeName, string attributeValue)
        {
            var attr = element.Attribute(attributeName);
            return attr != null && attr.Value == attributeValue;
        }

        /// <summary>
        /// Возвращает значение, указывающее, есть ли в этом элементе атрибут с заданным именем и, если есть, равно ли его значение заданному.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="attributeName">Имя атрибута.</param>
        /// <param name="attributeValue">Значение атрибута.</param>
        /// <returns>true если атрибут с таким именем есть и его значение равно заданному; иначе false</returns>
        public static bool HasAttributeValue(this XElement element, XName attributeName, int attributeValue)
        {
            var attr = element.Attribute(attributeName);
            return attr != null && attr.Value == attributeValue.ToString();
        }

        public static string[] Split(this string s, string separator)
        {
            return s.Split(new string[] { separator }, StringSplitOptions.None);
        }

        public static int GetDepth(this XElement element)
        {
            int depth = 0;
            while (element != null)
            {
                depth++;
                element = element.Parent;
            }
            return depth;
        }
    }

    static class Xtensions
    {
        public static void RemoveAttribute(this XElement elem, string name)
        {
            var attr = elem.Attribute(name);
            if (attr != null)
            {
                attr.Remove();
            }
        }

        #region HTML-экстеншены

        #region Работа со стилями
        public static Dictionary<string, string> GetStyles(this XElement elem)
        {
            var style = elem.Attribute("style");
            if (style == null)
            {
                return new Dictionary<string, string>();
            }

            var result = new Dictionary<string, string>();
            foreach (var styleAttr in style.Value.Split(';'))
            {
                var nameValue = styleAttr.Trim().Split(':');
                result.Add(nameValue[0].Trim(), nameValue[1].Trim());
            }

            return result;
        }

        public static bool HasStyle(this XElement elem, string name)
        {
            var style = elem.Attribute("style");
            if (style == null)
            {
                return false;
            }

            var result = new Dictionary<string, string>();
            foreach (var styleAttr in style.Value.Split(';'))
            {
                var nameValue = styleAttr.Trim().Split(':');
                if (nameValue[0].Trim() == name)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Возвращает значение указанного стиля в атрибуте style. Если такого стиля нет, возвращает null.
        /// </summary>
        /// <param name="elem"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetStyle(this XElement elem, string name)
        {
            var style = elem.Attribute("style");
            if (style == null)
            {
                return null;
            }

            var result = new Dictionary<string, string>();
            foreach (var styleAttr in style.Value.Split(';'))
            {
                var nameValue = styleAttr.Trim().Split(':');
                if (nameValue[0].Trim() == name)
                {
                    return nameValue[1].Trim();
                }
            }

            return null;
        }

        public static void RemoveStyle(this XElement elem, string name)
        {
            var styleAttribute = elem.Attribute("style");
            if (styleAttribute == null)
            {
                return;
            }

            var styles = styleAttribute.Value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            foreach (var style in styles.ToList())
            {
                var nameValue = style.Split(':');
                if (nameValue[0].Trim() == name)
                {
                    styles.Remove(style);
                }

            }

            styleAttribute.Value = string.Join("; ", styles);
        }

        /// <summary>
        /// Добавляет в атрибут style элемента текст в формате "name: value". Если атрибута нет, создаёт его.
        /// </summary>
        /// <param name="elem"></param>
        /// <param name="name">Название атрибута CSS-стиля.</param>
        /// <param name="value">Значение атрибута CSS-стиля.</param>
        public static XElement SetStyle(this XElement elem, string name, string value)
        {
            name = name.Trim();
            value = value.Trim();

            var styleAttribute = elem.Attribute("style");
            if (styleAttribute == null)
            {
                elem.Add(new XAttribute("style", name + ": " + value));
                return elem;
            }

            var styles = styleAttribute.Value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            if (styles.Any(p => p.Split(':')[0].Trim() == name))
            {
                for (int i = 0; i < styles.Count; i++)
                {
                    var nameValue = styles[i].Split(':');
                    if (nameValue[0].Trim() == name)
                    {
                        styles[i] = styles[i].Replace(nameValue[1].Trim(), value);
                    }
                }
            }
            else
            {
                styles.Add(string.Format("{0}: {1}", name, value));
            }

            styleAttribute.Value = string.Join("; ", styles);
            return elem;
        }

        public static void RemoveStyles(this XElement elem)
        {
            var styleAttribute = elem.Attribute("style");
            if (styleAttribute != null)
            {
                styleAttribute.Remove();
            }
        }
        #endregion

        public static int GetRowspan(this XElement elem, int defaultValue)
        {
            if (elem.HasAttribute("rowspan") && int.TryParse(elem.Attribute("rowspan").Value, out defaultValue))
            {
                return defaultValue;
            }

            return defaultValue;
        }

        public static int GetColspan(this XElement elem, int defaultValue)
        {
            if (elem.HasAttribute("colspan") && int.TryParse(elem.Attribute("colspan").Value, out defaultValue))
            {
                return defaultValue;
            }

            return defaultValue;
        }

        public static void AddClass(this XElement elem, string value)
        {
            var cl = elem.Attribute("class");
            if (cl != null)
            {
                cl.Value += " " + value;
            }
            else
            {
                elem.Add(new XAttribute("class", value));
            }
        }

        public static bool HasClass(this XElement elem, string value)
        {
            var classAttr = elem.Attribute("class");

            return classAttr != null && classAttr.Value == value;
        }
        #endregion
    }
}
