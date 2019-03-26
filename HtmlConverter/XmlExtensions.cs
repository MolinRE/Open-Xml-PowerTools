using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace HtmlConverter
{
    public class XRow : List<XElement>
    {
        public new XElement this[int index]
        {
            get
            {
                if (Count > index)
                {
                    return base[index];
                }

                return null;
            }
            set
            {
                if (Count <= index)
                {
                    // while (Count <= index)
                    for (int i = Count - 1; i < index; i++)
                    {
                        Add(null);
                    }
                }

                base[index] = value;
            }
        }

        public XElement Element
        {
            get { return this.Any() ? this.First().Parent : null; }
        }
    }
    
    public class XTable : List<XRow>
    {
        public new XRow this[int index]
        {
            get
            {
                if (Count > index)
                {
                    return base[index];
                }

                return null;
            }
            set
            {
                if (Count <= index)
                {
                    // while (Count <= index)
                    for (int i = Count - 1; i < index; i++)
                    {
                        Add(null);
                    }
                }

                base[index] = value;
            }
        }
    }

    static class Xtensions
    {
        public static bool HasAttribute(this XElement elem, XName name)
        {
            return elem.Attribute(name) != null;
        }

        public static bool HasAttributeValue(this XElement elem, XName name, string value)
        {
            return elem.Attribute(name)?.Value == value;
        }

        public static void RemoveAttribute(this XElement elem, string name)
        {
            var attr = elem.Attribute(name);
            if (attr != null)
            {
                attr.Remove();
            }
        }

        public static string GetValue(this XElement elem)
        {
            return string.Join("",
            elem.DescendantNodes()
                .Where(p => p.NodeType == XmlNodeType.Text || (p.NodeType == XmlNodeType.Element && ((XElement)p).Name.LocalName == "br"))
                .Select(s => s.NodeType == XmlNodeType.Text ? s.ToString().Trim() : "\n"));
        }

        public static bool HasChild(this XElement elem, string localName)
        {
            return elem.Elements().Any(p => p.Name.LocalName.Equals(localName));
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
            foreach (var styleAttr in style.Value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
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

    static class IEnumerableXtensions
    {
        /// <summary>
        /// Добавляет в атрибут style каждого элемента текст в формате "name: value". Если атрибута нет, создаёт его.
        /// </summary>
        /// <param name="elem"></param>
        /// <param name="name">Название атрибута CSS-стиля.</param>
        /// <param name="value">Значение атрибута CSS-стиля.</param>
        public static void SetStyle(this IEnumerable<XElement> elements, string name, string value)
        {
            foreach (var elem in elements)
            {
                name = name.Trim();
                value = value.Trim();

                var styleAttribute = elem.Attribute("style");
                if (styleAttribute == null)
                {
                    elem.Add(new XAttribute("style", name + ": " + value));
                    return;
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
            }
        }
    }

    /// <summary>
    /// Предикаты для поиска элементов
    /// </summary>
    static class X
    {
        /// <summary>
        /// Элементы типа параграф
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static bool Paragraph(XElement element)
        {
            return element.Name.LocalName == "p";
        }

        /// <summary>
        /// Элементы типа таблица
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static bool Table(XElement element)
        {
            return element.Name.LocalName == "table";
        }
    }
}
