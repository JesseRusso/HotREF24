using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        public static void CreateHRV()
        {

        }
    }
}
