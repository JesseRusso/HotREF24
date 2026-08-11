using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotPort
{
    internal class Ceiling
    {
        private string ceilingName;
        private string lengthMetric;
        private string areaMetric;
        private string ceilingType;
        private string typeCode;
        private string typeEng;
        private string typeFr;
        private string heelHeight;
        private string ceilingSlope;
        private string slopeCode;
        private string slopeValue;
        private string slopeEng;
        private string slopeFr;
        private string vaultRise;
        private string slopeName = "";
        private string rValue;
        private bool vaultCheck = false;
        private int id;

        public Ceiling(string name, string type, double area, double length, string slope, double heel)
        {
            ceilingName = name;
            ceilingType = type;
            ceilingSlope = slope;
            heelHeight = Math.Round(heel * 0.3048, 3).ToString();
            areaMetric = Math.Round(area * 0.092903, 4).ToString();
            lengthMetric = Math.Round(length * 0.3048, 4).ToString();
            SetType();
            SetSlope();
        }

        public Ceiling(object[] args, Dictionary<string,string> rValues, string builder)
        {
            ceilingName = (string)args[0];
            ceilingType = (string)args[1];
            areaMetric = Math.Round((double)args[2] * 0.092903, 4).ToString();
            lengthMetric = Math.Round((double)args[3] * 0.3048, 4).ToString();
            ceilingSlope = (string)args[4];
            heelHeight = Math.Round((double)args[5] * 0.3048, 3).ToString();
            id = (int)args[6];
            SetType();
            SetSlope();
            SetRValue(rValues, builder);
        }

        public Ceiling(object[] args, Dictionary<string, string> rValues, string builder, bool vault)
        {
            vaultCheck = vault;
            ceilingName = (string)args[0];
            ceilingType = (string)args[1];
            areaMetric = Math.Round((double)args[2] * 0.092903, 4).ToString();
            lengthMetric = Math.Round((double)args[3] * 0.3048, 4).ToString();
            ceilingSlope = (string)args[4];
            heelHeight = Math.Round((double)args[5] * 0.3048, 3).ToString();
            vaultRise = (string)args[6];
            id = (int)args[7];
            SetType();
            SetSlope();
            SetRValue(rValues, builder);
        }

        private void SetType()
        {
            switch (ceilingType.ToLower())
            {
                case "gable":
                    typeCode = "2";
                    typeEng = "Attic/gable";
                    typeFr = "Combles/pignon";
                    break;
                case "hip":
                    typeCode = "3";
                    typeEng = "Attic/hip";
                    typeFr = "Combles/arête";
                    break;
                case "cathedral":
                    typeCode = "4";
                    typeEng = "Cathedral";
                    typeFr = "Cathédrale";
                    break;
                case "flat":
                    typeCode = "5";
                    typeEng = "Flat";
                    typeFr = "Plat";
                    slopeName = "Flat";
                    ceilingSlope = "0";
                    break;
                case "scissor":
                    typeCode = "6";
                    typeEng = "Scissor";
                    typeFr = "Ciseaux";
                    break;
                default:
                    typeCode = "3";
                    typeEng = "Attic/hip";
                    typeFr = "Combles/arête";
                    break;
            }
        }

        private void SetRValue(Dictionary<string, string> values, string builder)
        {
            switch(ceilingType.ToLower())
            {
                case "gable":
                    rValue = values["ceiling"];
                    break;
                case "hip":
                    rValue = values["ceiling"];
                    break;
                case "cathedral":
                    rValue = values["cathedral"];
                    break;
                case "flat":
                    rValue = values["flat"];
                    break;
                case "scissor":
                    rValue = values["vault"];
                    break;
                default:
                    rValue = values["ceiling"];
                    break;
            }
            if (builder.ToLower().Contains("mckee") && Convert.ToDouble(vaultRise) < 5)
            {
                rValue = values["ceiling"];
            }
        }
        private void SetSlope()
        {
            switch (ceilingSlope)
            {
                case "0":
                    slopeCode = "1";
                    slopeValue = "0";
                    slopeEng = "Flat roof";
                    slopeFr = "Toit plat";
                    break;
                case "2":
                    slopeCode = "2";
                    slopeValue = "0.167";
                    slopeEng = "2 / 12";
                    slopeFr = "2 / 12";
                    break;
                case "3":
                    slopeCode = "3";
                    slopeValue = "0.25";
                    slopeEng = "3 / 12";
                    slopeFr = "3 / 12";
                    break;
                case "4":
                    slopeCode = "4";
                    slopeValue = "0.333";
                    slopeEng = "4 / 12";
                    slopeFr = "4 / 12";
                    break;
                case "5":
                    slopeCode = "5";
                    slopeValue = "0.417";
                    slopeEng = "5 / 12";
                    slopeFr = "5 / 12";
                    break;
                case "6":
                    slopeCode = "6";
                    slopeValue = "0.5";
                    slopeEng = "6 / 12";
                    slopeFr = "6 / 12";
                    break;
                case "7":
                    slopeCode = "7";
                    slopeValue = "0.583";
                    slopeEng = "7 / 12";
                    slopeFr = "7 / 12";
                    break;
                default:
                    slopeCode = "0";
                    slopeValue = Math.Round(System.Convert.ToDouble(ceilingSlope) / 12, 4).ToString();
                    slopeEng = "User specified";
                    slopeFr = "Spécifié par l'utilisateur";
                    break;
            }
            if (System.Convert.ToDouble(ceilingSlope) > 7)
            {
                slopeCode = "0";
                slopeValue = Math.Round(System.Convert.ToDouble(ceilingSlope)/12,4).ToString();
                slopeEng = "User specified";
                slopeFr = "Spécifié par l'utilisateur";
            }
        }
        public void AddCodeCeiling(XDocument house)
        {
            if (slopeCode.Equals("1"))
            {
                slopeName = "Flat";
            }
            else if (vaultCheck)
            {
                slopeName = "";
            
            }
            else slopeName = ceilingSlope + "/12";

            XElement comp = (from el in house.Descendants("Components")
                                 select el).First();
            comp.Add(
                new XElement("Ceiling",
                new XAttribute("id", id),
                    new XElement("Label", ceilingName + " " + slopeName),
                    new XElement("Construction",
                        new XElement("Type",
                            new XAttribute("code", typeCode),
                            new XElement("English", typeEng),
                            new XElement("French", typeFr)),
                        new XElement("CeilingType", "User specified",
                            new XAttribute("rValue", rValue),
                            new XAttribute("nominalInsulation", rValue))),
                    new XElement("Measurements",
                        new XAttribute("length", lengthMetric),
                        new XAttribute("area", areaMetric),
                        new XAttribute("heelHeight", heelHeight),
                            new XElement("Slope",
                                new XAttribute("code", slopeCode),
                                new XAttribute("value", slopeValue),
                                    new XElement("English", slopeEng),
                                    new XElement("French", slopeFr)))));
        }
        public void AddErsCeiling(XDocument house)
        {
            if (slopeCode.Equals("1"))
            {
                slopeName = "Flat";
            }
            else if (vaultCheck)
            {
                slopeName = "";

            }
            else slopeName = ceilingSlope + "/12";

            XElement comp = (from el in house.Descendants("Components")
                             select el).First();
            comp.Add(
                new XElement("Ceiling",
                new XAttribute("id", id),
                    new XElement("Label", ceilingName + " " + slopeName),
                    new XElement("Construction",
                        new XElement("Type",
                            new XAttribute("code", typeCode),
                            new XElement("English", typeEng),
                            new XElement("French", typeFr)),
                        new XElement("CeilingType", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("length", lengthMetric),
                        new XAttribute("area", areaMetric),
                        new XAttribute("heelHeight", heelHeight),
                            new XElement("Slope",
                                new XAttribute("code", slopeCode),
                                new XAttribute("value", slopeValue),
                                    new XElement("English", slopeEng),
                                    new XElement("French", slopeFr)))));
        }
    }
}
