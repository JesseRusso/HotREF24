using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotPort.Models
{
    internal class SolarPanel
    {
        public double Area { get; init; }
        public double Slope { get; init; }
        public double Azimuth { get; init; }
        public double Efficiency { get; init; }
        public double OperatingTemp { get; init; }
        public double TempCoeff { get; init; }
        public string Manufacturer { get; init; } = string.Empty;
        public string Model {  get; init; } = string.Empty;
        public int Rank { get; init; }

        public XElement NewPanel()
        {
            XElement panel = new XElement("System",
                                new XAttribute("rank", Rank),
                                new XElement("EquipmentInformation",
                                    new XElement("Manufacturer", Manufacturer),
                                    new XElement("Model", Model)),
                                new XElement("Array",
                                    new XAttribute("area", SquareFtToMetres()),
                                    new XAttribute("slope", Slope),
                                    new XAttribute("azimuth", Math.Abs(180 - Azimuth))),
                                new XElement("Efficiency",
                                    new XAttribute("miscellaneousLosses", "5"),
                                    new XAttribute("otherPowerLosses", "0"),
                                    new XAttribute("inverterEfficiency", "95"),
                                    new XAttribute("gridAbsorptionRate", "100")),
                                new XElement("Module",
                                    new XAttribute("efficiency", Efficiency),
                                    new XAttribute("cellTemperature", TempCelcius()),
                                    new XAttribute("coefficientOfEfficiency", MetricCoefficient()),
                                    new XElement("Type",
                                        new XAttribute("code", "6"),
                                        new XElement("English", "User Specified"),
                                        new XElement("French", "Spécifié par l'utilisateur"))),
                                new XElement("Orientation",
                                    new XAttribute("solarPanel", Azimuth),
                                    new XAttribute("degrees", "0"),
                                    new XAttribute("minutes", "0"),
                                    new XElement("MagGeo",
                                        new XAttribute("code", "0")),
                                    new XElement("Declination",
                                        new XAttribute("code", "1"),
                                            new XElement("English", "Westerly"),
                                            new XElement("French", "Vers l'ouest"))));
            return panel;

        }

        private double TempCelcius()
        {
            double celcius = Math.Round((OperatingTemp - 32) / 1.8, 4);
            return celcius;
        }
        private double MetricCoefficient()
        {
            return Math.Round(TempCoeff * 1.8, 4);
        }
        private double SquareFtToMetres()
        {
            return Math.Round(Area * 0.092903, 4);
        }
    }
}
