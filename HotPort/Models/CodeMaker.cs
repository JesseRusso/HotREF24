using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotPort.Models
{
    internal static class CodeMaker
    {
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
    }
}
