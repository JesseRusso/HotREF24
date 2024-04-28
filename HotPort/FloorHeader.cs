using System;
using System.Xml.Linq;

namespace HotPort
{
    internal static class FloorHeader
    {

        public static XElement NewJoist(string height, string rsi, string length, string id)
        {
            string Height = Math.Round(Convert.ToDouble(height) * 0.3048, 3).ToString();
            string RSI = rsi;
            string Length = Math.Round(Convert.ToDouble(length) * 0.3048, 3).ToString();
            string ID = id;

            XElement rimJoist = new XElement("FloorHeader",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", ID),
                new XElement("Label", "Rim Joist"),
                new XElement("Construction",
                    new XElement("Type", "User specified",
                        new XAttribute("rValue", RSI),
                        new XAttribute("nominalInsulation", "2.8507"))),
                new XElement("Measurements",
                    new XAttribute("height", Height),
                    new XAttribute("perimeter", Length)),
                new XElement("FacingDirection",
                    new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O")));
            return rimJoist;
        }
    }
}
