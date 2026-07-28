using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2010.Excel;
using HotPort.Models;
using System;
using System.Xml.Linq;

namespace HotPort
{
    internal static class FloorHeader
    {
        public static XElement NewJoist(string height, string rsi, string length, string id)
        {
            string _height = Math.Round(Convert.ToDouble(height) * 0.3048, 3).ToString();
            string _rsi = rsi;
            string _length = Math.Round(Convert.ToDouble(length) * 0.3048, 3).ToString();
            string _id = id;

            XElement rimJoist = new XElement("FloorHeader",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", _id),
                new XElement("Label", "Rim Joist"),
                new XElement("Construction",
                    new XElement("Type", "User specified",
                        new XAttribute("rValue", _rsi),
                        new XAttribute("nominalInsulation", "2.8507"))),
                new XElement("Measurements",
                    new XAttribute("height", _height),
                    new XAttribute("perimeter", _length)),
                new XElement("FacingDirection",
                    new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O")));
            return rimJoist;
        }

        public static XElement NewErsJoist(string height, string length, string id)
        {
            string _id = id;

            XElement rimJoist = new XElement("FloorHeader",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", _id),
                new XElement("Label", "Rim Joist"),
                new XElement("Construction",
                    new XElement("Type", "",
                        new XAttribute("idref", ""),
                        new XAttribute("rValue", "0"),
                        new XAttribute("nominalInsulation", "0"))),
                new XElement("Measurements",
                    new XAttribute("height", height),
                    new XAttribute("perimeter", length)),
                new XElement("FacingDirection",
                    new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O")));

            return rimJoist;
        }
    }
}
