using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotPort
{
    internal static class CodeTools
    {
        /**
         * <param name="codeID">The codeID to assign to this ccode</param>
         * <param name="rsi">The RSI value of the window for this code</param>
         * <param name="shgc">SHGC of the window code</param>
         */
        public static XElement StandardWindowCode(double rsi, double shgc, int codeID)
        {
            XElement code =
                new XElement("Code",
                    new XAttribute("id", codeID),
                    new XAttribute("value", "100000"),
                    new XAttribute("nomimalRValue", "0"),
                    new XElement("Label", "100000"),
                    new XElement("Description", "100000"),
                    new XElement("Layers",
                        new XElement("GlazingTypes",
                            new XAttribute("code", "1")),
                        new XElement("CoatingsTints",
                            new XAttribute("code", "0")),
                        new XElement("Type",
                            new XAttribute("code", "0")),
                        new XElement("FrameMaterial",
                            new XAttribute("code", "1"))));
            return code;
        }
        public static XElement AddWindowCode(Window window, XDocument house)
        {
            string codeString = "Code " + GetValidCodeID(house);

            XElement code = new XElement("Code",
                new XAttribute("id", codeString),
                new XAttribute("nominalRValue", "0"),
                    new XElement("Label", window.ToString()),
                    new XElement("Description", String.Empty),
                    new XElement("Layers",
                        new XElement("WindowLegacy",
                        new XAttribute("frameHeight", "0"),
                        new XAttribute("shgc", "0.5"),
                        new XAttribute("rank", "1"),
                        new XElement("Type",
                            new XAttribute("code", "1"),
                            new XElement("English", "Picture"),
                            new XElement("French", "Fixe")),
                        new XElement("RsiValues",
                            new XAttribute("centreOfGlass", window.RSI),
                            new XAttribute("edgeOfGlass", window.RSI),
                            new XAttribute("frame", window.RSI)))));
            return code;
        }
        /**
         * Returns the next valid code ID for the house file
         */
        public static int GetValidCodeID(XDocument house)
        {
            List<string> codeIDs = new List<string>();
            var hasCode = from el in house.Descendants("Codes").Descendants().Attributes("id")
                          select el.Value;

            foreach (string code in hasCode)
            {
                string[] codeStrings = code.Split(' ');
                codeIDs.Add(codeStrings[1]);
            }
            codeIDs.Sort();
            return int.Parse(codeIDs.Last()) + 1;
        }
        public static int GetMaxWindowCodeID(XDocument house)
        {
            XElement? codesBLock = house?.Root?.Element("Codes");
            IEnumerable<XElement>? codes = codesBLock?.Descendants("Code");
            IEnumerable<XElement>? windowCodes = codes?.Descendants("Window");

            if (windowCodes != null && windowCodes.Any())
            {
                return int.Parse(windowCodes.First().Attribute("id").Value);
            }

            return codes.Count() + 1;
        }
        /**
         * Searches the <Codes></Codes> block for existing window codes that match the one specified by the window
         * Adds the code to the <Codes></Codes> block if it doesn't exist
         */
        public static int FindWindowCode(XDocument house, Window window)
        {
            XElement? windowCodes = house?.Root?.Element("Codes")?.Element("Window");
            XElement? userBlock = windowCodes?.Element("UserDefined");
            List<XElement>? standardCodes = windowCodes?.Element("Standard")?.Descendants("Code").ToList();
            List<XElement>? userDefiend = userBlock?.Descendants("Code").ToList();

            if(userBlock != null && userDefiend.Any())
            {
                foreach (XElement code in userDefiend)
                {
                    if (window.ToString() == code.Element("Label")?.Value.ToString())
                    {
                        string[] idString = code.Attribute("id").Value.ToString().Split(' ');
                        int id = int.Parse(idString[1]);
                        return id;
                    }
                }
            }
            if(userBlock == null)
            {
                XElement block = new XElement("UserDefined");
                block.Add(AddWindowCode(window, house));
                windowCodes.Add(block);
            }
            else
            {
                userBlock.Add(AddWindowCode(window, house));
            }
            return GetValidCodeID(house) - 1;
        }
    }
}
