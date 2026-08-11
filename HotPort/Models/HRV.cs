using System;
using System.Xml.Linq;

namespace HotPort
{
    internal static class HRV
    {
        public static XElement CreateFan(double flowrate, string fanPower)
        {
            XElement vent = new XElement("BaseVentilator",
                        new XAttribute("supplyFlowrate", Math.Round(flowrate * 0.471947, 4).ToString()),
                        new XAttribute("exhaustFlowrate", Math.Round(flowrate * 0.471947, 4).ToString()),
                        new XAttribute("fanPower1", fanPower),
                        new XAttribute("isDefaultFanpower", "false"),
                        new XAttribute("isEnergyStar", "false"),
                        new XAttribute("isHomeVentilatingInstituteCertified", "false"),
                        new XAttribute("isSupplemental", "false"),
                            new XElement("EquipmentInformation"),
                            new XElement("VentilatorType",
                            new XAttribute("code", "4"),
                                new XElement("English", "Utility"),
                                new XElement("French", "Utilité")));
            return vent;
        }
        public static XElement CreateHRV()
        {
            XElement Duct(string name) =>
                new XElement(name,
                    new XAttribute("length", "1.5"),
                    new XAttribute("diameter", "152.4"),
                    new XAttribute("insulation", "0.7"),
                    new XElement("Location",
                        new XAttribute("code", "4"),
                        new XElement("English", "Main floor"),
                        new XElement("French", "Rez-de-chaussée")),
                    new XElement("Type",
                        new XAttribute("code", "1"),
                        new XElement("English", "Flexible"),
                        new XElement("French", "Flexible")),
                    new XElement("Sealing",
                        new XAttribute("code", "2"),
                        new XElement("English", "Sealed"),
                        new XElement("French", "Scellé")));

            return new XElement("Hrv",
                new XAttribute("supplyFlowrate", "27.8449"),
                new XAttribute("exhaustFlowrate", "27.8449"),
                new XAttribute("fanPower1", "88"),
                new XAttribute("isDefaultFanpower", "false"),
                new XAttribute("isEnergyStar", "false"),
                new XAttribute("isHomeVentilatingInstituteCertified", "false"),
                new XAttribute("isSupplemental", "false"),
                new XAttribute("temperatureCondition1", "0"),
                new XAttribute("temperatureCondition2", "-25"),
                new XAttribute("fanPower2", "85"),
                new XAttribute("efficiency1", "80"),
                new XAttribute("efficiency2", "62"),
                new XAttribute("preheaterCapacity", "0"),
                new XAttribute("lowTempVentReduction", "0"),
                new XAttribute("coolingEfficiency", "25"),
                new XElement("EquipmentInformation",
                    new XElement("Manufacturer", "Fantech"),
                    new XElement("Model", "HERO 200H")),
                new XElement("VentilatorType",
                    new XAttribute("code", "1"),
                    new XElement("English", "HRV/ERV"),
                    new XElement("French", "VRC/VRE")),
                new XElement("ColdAirDucts",
                    Duct("Supply"),
                    Duct("Exhaust")));
        }
    }
}
