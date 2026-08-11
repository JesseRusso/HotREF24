using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotPort.Models
{
    internal static class Fireplace
    {
        public static XElement NewFireplace()
        {
            return new XElement("SupplementaryHeatingSystems",
                new XElement("System",
                    new XAttribute("rank", "1"),
                    new XElement("EquipmentInformation",
                        new XAttribute("csaEpa", "false")),
                    new XElement("Equipment",
                        new XElement("EnergySource",
                            new XAttribute("code", "2"),
                            new XElement("English", "Natural gas"),
                            new XElement("French", "Gaz naturel")),
                        new XElement("Type",
                            new XAttribute("code", "3"),
                            new XElement("English", "Fireplace with spark ignit. (sealed)"),
                            new XElement("French", "Foyer avec allum. par étincelle (scellé)"))),
                    new XElement("Specifications",
                        new XAttribute("efficiency", "30"),
                        new XAttribute("pilotLight", "0"),
                        new XAttribute("damperClosed", "false"),
                        new XElement("YearMade",
                            new XAttribute("code", "10"),
                            new XElement("English", "2000-"),
                            new XElement("French", "2000-")),
                        new XElement("Usage",
                            new XAttribute("code", "1"),
                            new XElement("English", "Never"),
                            new XElement("French", "Jamais")),
                        new XElement("LocationHeated",
                            new XAttribute("code", "1"),
                            new XAttribute("value", "12.0031"),
                            new XElement("English", "Main Floors"),
                            new XElement("French", "Plancher Principaux")),
                        new XElement("Flue",
                            new XAttribute("isInterior", "true"),
                            new XAttribute("diameter", "0"),
                            new XElement("Type",
                                new XAttribute("code", "1"),
                                new XElement("English", "Brick"),
                                new XElement("French", "Brique"))),
                        new XElement("OutputCapacity",
                            new XAttribute("value", "2"),
                            new XAttribute("uiUnits", "kW")))));
        }
    }
}
