//Created by Jesse Russo 2021
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotPort
{
    public class CreateProp
    {
        private static string filePath;
        public static XDocument newHouse;
        public static string builder;
        public static int maxID;
        public string? wallRValue;
        public string? floorRValue;
        public static string ceilingRValue;
        public static string vaultRValue;
        public static string cathedralRValue;
        public static string flatCeilingRValue;
        public static int ceilingCount = 1;
        private static bool basementPresent = true;
        private static int maxWindowRow = 62;


        public CreateProp(string excelFilePath, XDocument template)
        {
            filePath = excelFilePath;
            newHouse = new XDocument(template);
            GetBuilder();
            CityCheck();
            FindID(template);
        }
        public static XDocument GetHouse()
        {
            return newHouse;
        }

        /**
         * Finds the largest component ID in the template and stores it. 
         * New components such as windows and doors are assigned new IDs that must not conflict with existing IDs
         */
        public static void FindID(XDocument newHouse)
        {
            List<string> ids = new();
            var hasID =
                from el in newHouse.Descendants("House").Descendants().Attributes("id")
                where el.Value != null
                select el.Value;

            foreach (string id in hasID)
            {
                ids.Add(id);
            }
            List<int> idList = ids.Select(s => int.Parse(s)).ToList();
            maxID = idList.Max() + 1;
            ids.Clear();
        }
        //Gets the builder name 
        private static void GetBuilder()
        {
            string builderName = (from el in newHouse.Descendants("File").Descendants("BuilderName")
                                  where el.IsEmpty == false
                                  select el).First().Value;
            builder = builderName;
        }
        //Checks special city cells P1:P5 in spreadsheet and sets
        //city-specific settings like weather station location,
        //NAI terrain info, and city name in specifications
        private void CityCheck()
        {
            string city;
            if (GetCellValue("Calc", "P1")?.ToLower() == "y")
            {
                city = "Okotoks";
            }
            else if (GetCellValue("Calc", "P2")?.ToLower() == "y")
            {
                city = "Edmonton";
            }
            else if (GetCellValue("Calc", "P3")?.ToLower() == "y")
            {
                city = "Airdrie";
            }
            else if (GetCellValue("Calc", "P4")?.ToLower() == "y")
            {
                city = "Victoria";
            }
            else if (GetCellValue("Calc", "P5")?.ToLower() == "y")
            {
                city = "Cochrane";
            }
            else { city = "Calgary"; }

            ChangeCityWeather(city);
        }
        private void ChangeCityWeather(string city)
        {
            newHouse.Element("HouseFile").Element("ProgramInformation").Element("Client").Element("StreetAddress").Element("City").SetValue(city);
            switch (city)
            {
                case "Cochrane":
                    newHouse.Element("HouseFile").Element("ProgramInformation").Element("Weather").Element("Location").SetAttributeValue("code", 10);
                    break;
                case "Okotoks":
                    newHouse.Element("HouseFile").Element("House").Element("NaturalAirInfiltration").Element("Specifications").Element("BuildingSite").Element("Terrain").SetAttributeValue("code", 3);
                    break;
            }
        }

        /* Copies R value from first wall element with "1st" in its name to be used in new wall creation.
         * Changes properties of the first wall and 2nd storey wall. Removes 2nd storey if 2nd storey area is 0 in spreadsheet
         * Changes Storeys property in house specifications if 2nd flr wall is removed */
        public void ChangeWalls()
        {
            XElement first = (XElement)(from el in newHouse.Descendants("Wall")
                                        where el.Element("Label").Value.Contains("1st")
                                        select el).First();

            wallRValue = first.Element("Construction").Element("Type").Attribute("rValue").Value.ToString();
            string? rimRValue = first.Element("Components")?.Element("FloorHeader")?.Element("Construction")?.Element("Type")?.Attribute("rValue")?.Value.ToString();

            first.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G21")) * 0.3048, 3));
            first.Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H21")) * 0.3048, 3));
            first.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E21"));
            first.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F21"));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "K21")) * 0.3048, 3));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "J21")) * 0.3048, 3));

            XElement second = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.Contains("2nd")
                                         select el).First();

            if (GetCellValue("Calc", "D4") == null || GetCellValue("Calc", "D4") == "0")
            {
                second.Remove();
                first.Descendants("FloorHeader").Remove();
                //change house to 1 storey in specifications if 2nd flr is not present
                foreach (XElement el in newHouse.Descendants("Specifications").Descendants("Storeys"))
                {
                    el.SetAttributeValue("code", "1");
                    el.Element("English").SetValue("One storey");
                    el.Element("French").SetValue("Un étage");
                }
            }
            else
            {
                second.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G24")) * 0.3048, 3));
                second.Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H24")) * 0.3048, 3));
                second.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E24"));
                second.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F24"));

                if(Convert.ToDouble(GetCellValue("Calc", "J24")) > 0)
                {
                    second.Element("Components")?.Add(FloorHeader.NewJoist(GetCellValue("Calc", "K24"), rimRValue, GetCellValue("Calc", "J24"), maxID.ToString()));
                    maxID++;
                }
            }
            GarageWall();
            TallWall();
            PlumbingWall();
        }

        //Finds garage wall in template and modifies its properties if it is present in the spreadsheet. Removes it from the file if it is absent.
        private void GarageWall()
        {
            XElement garage = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.ToLower().Contains("garage")
                                         select el)?.First();

            if (GetCellValue("Calc", "K3") == null)
            {
                garage.Remove();
            }
            else
            {
                garage.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G30")) * 0.3048, 3));
                garage.Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H30")) * 0.3048, 3));
                garage.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E30"));
                garage.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F30"));
                garage.Element("Components").Element("FloorHeader")?.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "K38")) * 0.3048, 3));
                garage.Element("Components").Element("FloorHeader")?.Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "J30")) * 0.3048, 3));
            }
        }

        //Finds tall wall in template and modifies its properties if it is present in the spreadsheet. Removes it from the file if it is absent.
        private void TallWall()
        {
            XElement tall = newHouse.Descendants("Wall").Where(x => x.Element("Label").Value.ToLower().Contains("tall")).First();
            double perim = Convert.ToDouble(GetCellValue("Calc", "H31"));

            if (GetCellValue("Calc", "K4") == null || perim <= 0)
            {
                tall.Remove();
            }
            else
            {
                tall.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G31")) * 0.3048, 3));
                tall.Element("Measurements").SetAttributeValue("perimeter", Math.Round(perim * 0.3048, 3));
                tall.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E31"));
                tall.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F31"));
            }
        }
        //Finds plumbing wall when label contains "plumbing" and modifies its properties if its present in spreadsheet. Removes it from the file if absent.
        private void PlumbingWall()
        {
            XElement plumbing = (XElement)(from el in newHouse.Descendants("Wall")
                                           where el.Element("Label").Value.ToLower().Contains("plumbing")
                                           select el)?.First();

            if (GetCellValue("Calc", "H34") == "0" && plumbing != null)
            {
                plumbing.Remove();
            }
            else
            {
                plumbing.Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G34")) * 0.3048, 3));
                plumbing.Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H34")) * 0.3048, 3));
                plumbing.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E34"));
                plumbing.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F34"));
            }

        }
        public void ChangeBasment()
        {
            bool basement = GetCellValue("Calc", "N2")?.ToLower() == "y" ? true : false;
            bool bsmtUnder4Feet = GetCellValue("Calc", "N4")?.ToLower() == "y" ? true : false;
            bool slabOnGrade = GetCellValue("Calc", "N5")?.ToLower() == "y" ? true : false;
            if(!basement && !bsmtUnder4Feet)
            {
                basementPresent = false;
            }

            if (basement)
            {
                XElement over4 = (XElement)(from el in newHouse.Descendants("Components").Descendants("Basement")
                                            where el.Element("Label").Value.Contains(">")
                                            select el).First();

                over4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "E38")) * 0.3048, 4));
                over4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "D38")) * 0.3048, 4));
                over4.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(Convert.ToDouble(GetCellValue("Calc", "F38")) * 0.092903, 4));
                over4.Element("Wall").Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G38")) * 0.3048, 4));
                over4.Element("Wall").Element("Measurements").SetAttributeValue("depth", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H38")) * 0.3048, 4));
                over4.Element("Wall").Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "J38"));
                over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "K38")) * 0.3048, 4));
                over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "L38")) * 0.3048, 4));

                if (Convert.ToDouble(GetCellValue("Calc", "I38")) > 1D || bsmtUnder4Feet)
                {
                    over4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                    over4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(Convert.ToDouble(GetCellValue("Calc", "I38")) * 0.3048, 4));
                    over4.Element("Wall").Element("Construction").Add(
                        new XElement("PonyWallType",
                            new XAttribute("nominalInsulation", "3.2536"),
                            new XElement("Description", "User specified"),
                            new XElement("Composite",
                                new XElement("Section",
                                    new XAttribute("rank", "1"),
                                    new XAttribute("percentage", "100"),
                                    new XAttribute("rsi", "2.6029"),
                                    new XAttribute("nominalRsi", "3.2536")))));
                }
                else
                {
                    over4.Element("Wall").SetAttributeValue("hasPonyWall", "false");
                    over4.Element("Wall").Element("Construction").Element("PonyWallType")?.Remove();
                    over4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", "0");
                }
                Under4Bsmt(bsmtUnder4Feet);
                SlabOnGrade(slabOnGrade);
                return;
            }
            XElement over4ft = newHouse.Descendants("Components").Descendants("Basement").Where(x => x.Element("Label").Value.Contains(">")).FirstOrDefault();
            over4ft?.Remove();
            Under4Bsmt(bsmtUnder4Feet);
            SlabOnGrade(slabOnGrade);
        }
        private void Under4Bsmt(bool under4Present)
        {
            bool check = under4Present;
            XElement under4 = (XElement)(from el in newHouse.Descendants("Components").Descendants("Basement")
                                         where el.Element("Label").Value.Contains("<")
                                         select el).First();
            if (check)
            {
                under4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "E39")) * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "D39")) * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(Convert.ToDouble(GetCellValue("Calc", "F39")) * 0.092903, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "G39")) * 0.3048, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("depth", Math.Round(Convert.ToDouble(GetCellValue("Calc", "H39")) * 0.3048, 4));
                under4.Element("Wall").Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "J39"));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(Convert.ToDouble(GetCellValue("Calc", "K39")) * 0.3048, 4));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "L39")) * 0.3048, 4));
                under4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                under4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(Convert.ToDouble(GetCellValue("Calc", "I39")) * 0.3048, 4));
            }
            else
            {
                under4.Remove();
            }
        }
        private void SlabOnGrade(bool isSlabPresent)
        {
            bool check = isSlabPresent;
            XElement slab = (XElement)(from el in newHouse.Descendants("Components").Descendants("Slab")
                                       where el.Element("Label").Value.Contains("Slab")
                                       select el).First();
            if (check)
            {
                slab.SetAttributeValue("exposedSurfacePerimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "E40")) * 0.3048, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(Convert.ToDouble(GetCellValue("Calc", "F40")) * 0.092903, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "D40")) * 0.3048, 4));
            }
            else
            {
                slab.Remove();
            }
        }

        // Gathers R values from existing ceiling elements in the H2K template
        public void CheckCeilings()
        {
            XElement ceil = (XElement)(from el in newHouse.Descendants("Ceiling")
                                       where el.Element("Label").Value.Contains("2nd")
                                       select el).First();

            ceilingRValue = ceil.Element("Construction").Element("CeilingType").Attribute("rValue").Value.ToString();
            ceil.Remove();

            XElement vault = (XElement)(from el in newHouse.Descendants("Components").Descendants("Ceiling")
                                        where el.Element("Construction").Element("Type").Attribute("code").Value == "6"
                                        select el).First();

            vaultRValue = vault.Element("Construction").Element("CeilingType").Attribute("rValue").Value;
            vault.Remove();

            XElement cath = (from el in newHouse.Descendants("Components").Descendants("Ceiling")
                             where (el.Element("Construction").Element("Type").Attribute("code").Value == "4")
                             where (el != null)
                             select el)?.FirstOrDefault();

            if (cath != null)
            {
                cathedralRValue = cath.Element("Construction").Element("CeilingType").Attribute("rValue").Value.ToString();
                cath.Remove();
            }
            else
            {
                cathedralRValue = vaultRValue;
            }

            XElement? flat = (XElement)(from el in newHouse?.Descendants("Components")?.Descendants("Ceiling")
                                       where (el.Element("Construction").Element("Type")?.Attribute("code").Value == "5")
                                       where (el != null)
                                       select el)?.FirstOrDefault();
            if (flat != null)
            {
                flatCeilingRValue = flat.Element("Construction").Element("CeilingType").Attribute("rValue").Value.ToString();
                flat.Remove();
            }
            else
            {
                flatCeilingRValue = vaultRValue;
            }
        }

        public void ChangeEquipment()
        {
            ChangeHRV();
            ChangeFurnace();
        }
        private void ChangeHRV()
        {
            string hrvMake = GetCellValue("Summary", "D74");
            string hrvModel = GetCellValue("Summary", "E74");
            string hrvPower1 = GetCellValue("General", "I4");
            string hrvPower2 = GetCellValue("General", "J4");
            string hrvSRE1 = GetCellValue("General", "I5");
            string hrvSRE2 = GetCellValue("General", "J5");
            double hrvFlowrate = Math.Round(Convert.ToDouble(GetCellValue("General", "H4")), 1);
            string fanPower = Math.Round(Convert.ToDouble(GetCellValue("General", "K4")), 1).ToString();

            if (Convert.ToDouble(GetCellValue("General", "I4")) <= 0)
            {
                foreach (XElement vent in newHouse.Descendants("WholeHouseVentilatorList"))
                {
                    vent.Element("Hrv")?.Remove();
                    vent.Add(HRV.CreateFan(hrvFlowrate, fanPower));
                }
            }
            else
            {
                foreach (XElement hrv in newHouse.Descendants("WholeHouseVentilatorList"))
                {
                    hrv.Element("Hrv").Element("EquipmentInformation").Element("Manufacturer").SetValue(hrvMake);
                    hrv.Element("Hrv").Element("EquipmentInformation").Element("Model").SetValue(hrvModel);
                    hrv.Element("Hrv").SetAttributeValue("supplyFlowrate", Math.Round(hrvFlowrate * 0.471947, 4).ToString());
                    hrv.Element("Hrv").SetAttributeValue("exhaustFlowrate", Math.Round(hrvFlowrate * 0.471947, 4).ToString());
                    hrv.Element("Hrv").SetAttributeValue("fanPower1", hrvPower1);
                    hrv.Element("Hrv").SetAttributeValue("fanPower2", hrvPower2);
                    hrv.Element("Hrv").SetAttributeValue("efficiency1", hrvSRE1);
                    hrv.Element("Hrv").SetAttributeValue("efficiency2", hrvSRE2);
                }
            }
        }
        private void ChangeFurnace()
        {
            string furnaceModel = GetCellValue("Summary", "B78");
            string furnaceOutput = GetCellValue("General", "C4");
            double btus = Convert.ToDouble(furnaceOutput);

            if (Convert.ToDouble(GetCellValue("General", "B4")) > 0)
            {
                furnaceModel += $" & {GetCellValue("Summary", "B79")}";
            }
            // Changes furnace output capacity and EF values
            foreach (XElement furn in newHouse.Descendants("Furnace"))
            {
                furn.Element("Specifications")?.SetAttributeValue("efficiency", GetCellValue("General", "A5"));
                furn.Element("Specifications")?.Element("OutputCapacity")?.SetAttributeValue("value", Math.Round(btus * 0.00029307107, 5).ToString());
                furn.Element("EquipmentInformation").Element("Manufacturer").SetValue(GetCellValue("Summary", "A78"));
                furn.Element("EquipmentInformation").Element("Model").SetValue(furnaceModel);
            }
            foreach (XElement fan in newHouse.Descendants("HeatingCooling"))
            {
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("low", GetCellValue("General", "E4"));
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("high", GetCellValue("General", "D4"));
            }
        }
        public void GasDHW()
        {
                foreach (XElement tank in newHouse.Descendants("Components").Descendants("HotWater"))
                {
                    tank.Element("Primary").DescendantsAndSelf().Remove();
                }
                bool isPrimary = true;
                if (Convert.ToDouble(GetCellValue("General", "P4")) > 0)
                {
                    string dhwMake = GetCellValue("Summary", "J74");
                    string dhwModel = GetCellValue("Summary", "K74");
                    string dhwSize = GetCellValue("Summary", "K75");
                    string dhwEF = GetCellValue("Summary", "K77");
                    bool isUEF = !GetCellValue("General", "P6").Equals("0");
                    string drawPattern = GetCellValue("General", "P6");
                    WaterHeater tank = new WaterHeater(dhwMake, dhwModel, dhwEF, dhwSize, false, isPrimary, isUEF, drawPattern);
                    tank.AddTank(basementPresent);
                }
        }
        public void ElectricDHW()
        {
            bool isPrimary = false;
            if (Convert.ToDouble(GetCellValue("General", "I32")) > 0)
            {
                if (Convert.ToDouble(GetCellValue("General", "P4")) <= 0)
                {
                    isPrimary = true;
                }
                string electricTankMake = GetCellValue("Summary", "L74");
                string electricTankModel = GetCellValue("Summary", "M74");
                string electricTankVolume = GetCellValue("Summary", "M75");
                string electricTankEF = GetCellValue("Summary", "M77");
                bool isUEF = false;
                string drawPattern = "none";

                WaterHeater tank = new WaterHeater(electricTankMake, electricTankModel, electricTankEF, electricTankVolume, true, isPrimary, isUEF, drawPattern);
                tank.AddTank(basementPresent);
            }
            }
        //Checks excel sheet for selected A/C unit and modifies file if there is a model chosen
        public void CheckAC()
        {
            string ACBtus = GetCellValue("General", "Q26");
            if (Convert.ToInt32(ACBtus) > 0)
            {
                XElement? bsmtTemps = newHouse?.Element("HouseFile")?.Element("House")?.Element("Temperatures")?.Element("Basement");
                bsmtTemps?.SetAttributeValue("cooled", "true");

                XElement? type2HeatingCooling = (from el in newHouse?.Descendants("HeatingCooling").Descendants("Type2")
                                                          select el).FirstOrDefault();

                type2HeatingCooling?.Add(
                    new XElement("AirConditioning",
                        new XElement("EquipmentInformation",
                            new XAttribute("energystar", "false"),
                            new XElement("Manufacturer", GetCellValue("General", "N25")),
                            new XElement("Model", GetCellValue("General", "P25"))),
                        new XElement("Equipment",
                            new XAttribute("crankcaseHeater", "60"),
                            new XElement("CentralType",
                                new XAttribute("code", "1"))),
                        new XElement("Specifications",
                            new XAttribute("sizingFactor", "1"),
                                new XElement("RatedCapacity",
                                    new XAttribute("code", "2"),
                                    new XAttribute("value", "2.015"),
                                    new XAttribute("uiUnits", "btu/hr")),
                            new XElement("Efficiency", new XAttribute("isCop", "false"), new XAttribute("value", GetCellValue("General", "O26")))),
                        new XElement("CoolingParameters", new XAttribute("sensibleHeatRatio", "0.76"), new XAttribute("openableWindowArea", "0"),
                            new XElement("FansAndPump", new XAttribute("flowRate", "123"), new XAttribute("hasEnergyEfficientMotor", "false"),
                                new XElement("Mode", new XAttribute("code", "1")),
                                new XElement("Power", new XAttribute("isCalculated", "true"))))));
            }
        }

        //Separates the strings in the filename and writes them as the address in the file
        public static void ChangeAddress(string address)
        {
            foreach (XElement add in newHouse.Descendants("Client").Descendants("StreetAddress"))
            {
                add.Element("Street")?.SetValue(CamelCaseToSpaceSeparated(address));
            }
        }

        //Changes values in specifications screen along with the house volume and highest ceiling height
        public void ChangeSpecs()
        {
            string secondFloor = GetCellValue("Calc", "E4");
            string aboveGrade = GetCellValue("Calc", "E6");
            string belowGrade = GetCellValue("Calc", "F6");
            string type = GetCellValue("Calc", "K5");

            foreach (XElement grade in newHouse.Descendants("Specifications").Descendants("HeatedFloorArea"))
            {
                double areaAbove = Convert.ToDouble(aboveGrade);
                double areaBelow = Convert.ToDouble(belowGrade);
                areaAbove = Math.Round(areaAbove * 0.092903, 1);
                areaBelow = Math.Round(areaBelow * 0.092903, 1);
                aboveGrade = areaAbove.ToString();
                belowGrade = areaBelow.ToString();

                grade.SetAttributeValue("aboveGrade", aboveGrade);
                grade.SetAttributeValue("belowGrade", belowGrade);
            }

            foreach (XElement infil in newHouse.Descendants("NaturalAirInfiltration"))
            {
                infil.Element("Specifications").Element("House").SetAttributeValue("volume", Math.Round(Convert.ToDouble(GetCellValue("Calc", "Q49")) * 0.02831684609, 4).ToString());
                infil.Element("Specifications").Element("BuildingSite").SetAttributeValue("highestCeiling", Math.Round(Convert.ToDouble(GetCellValue("Calc", "M51")) * 0.3048, 4).ToString());
            }

            //This section gets the number of corners from each floor and finds the maximum value. Then sets the plan shape 
            List<string> corners = new List<string>();
            corners.Add(GetCellValue("Calc", "F2"));
            corners.Add(GetCellValue("Calc", "F3"));
            corners.Add(GetCellValue("Calc", "F4"));

            corners.RemoveAll(Value => Value == null);
            List<int> shape = new List<int>(corners.Select(s => int.Parse(s)).ToList());
            int maxCorners = shape.Max();
            corners.Clear();

            foreach (XElement ps in newHouse.Descendants("PlanShape"))
            {
                if (maxCorners <= 4)
                {
                    ps.SetAttributeValue("code", "1");
                    ps.Element("English")?.SetValue("Rectangular");
                    ps.Element("French")?.SetValue("Rectangulaire");
                }
                if (maxCorners > 4 && maxCorners < 7)
                {
                    ps.SetAttributeValue("code", "4");
                    ps.Element("English")?.SetValue("Other, 5-6 corners");
                    ps.Element("French")?.SetValue("Autre, 5-6 coins");
                }
                if (maxCorners > 6 && maxCorners < 9)
                {
                    ps.SetAttributeValue("code", "5");
                    ps.Element("English")?.SetValue("Other, 7-8 corners");
                    ps.Element("French")?.SetValue("Autre, 7-8 coins");
                }
                if (maxCorners > 8 && maxCorners < 11)
                {
                    ps.SetAttributeValue("code", "6");
                    ps.Element("English")?.SetValue("Other, 9-10 corners");
                    ps.Element("French")?.SetValue("Autre, 9-10 coins");
                }
                if (maxCorners >= 11)
                {
                    ps.SetAttributeValue("code", "7");
                    ps.Element("English")?.SetValue("Other, 11 or more corners");
                    ps.Element("French")?.SetValue("Autre, 11 coins ou plus");
                }
            }
            foreach (XElement houseType in newHouse.Descendants("HouseType"))
            {
                string partyFirstFlr = GetCellValue("Calc", "H3");
                string? calcK5 = GetCellValue("Calc", "K5")?.ToLower();
                if (double.TryParse(partyFirstFlr, out double value) && value > 0)
                {
                    switch (calcK5) 
                    {
                        case "single":
                            break;
                        case "semi":
                            houseType.SetAttributeValue("code", "2");
                            break;
                        case "rowhouse-end":
                            houseType.SetAttributeValue("code", "6");
                            break;
                        case "rowhouse-mid":
                            houseType.SetAttributeValue("code", "8");
                            break;
                        case "y":
                            houseType.SetAttributeValue("code", "2");
                            break;
                    }
                }
            }
            SetDate();
        }
        public static void SetDate()
        {
            XElement evalDate = newHouse.Element("HouseFile").Element("ProgramInformation").Element("File");
            DateTime dateTime = DateTime.UtcNow;
            dateTime = dateTime.Date;
            string date = dateTime.ToString("yyyy-MM-dd");
            evalDate.SetAttributeValue("evaluationDate", date);
        }
        public void ChangeFloors()
        {
            double garFlrArea = Convert.ToDouble(GetCellValue("Calc", "P21"));
            double garFlrLength = Convert.ToDouble(GetCellValue("Calc", "O21"));

            XElement garFlr = (XElement)(from el in newHouse.Descendants("Floor")
                                         where el.Element("Label").Value.ToLower().Contains("garage")
                                         select el).First();

            XElement floor = (XElement)(from el in newHouse.Descendants("Floor")
                                        where el.Element("Label").Value.ToLower().Contains("cant")
                                        select el).First();

            floorRValue = floor.Element("Construction")?.Element("Type")?.Attribute("rValue")?.Value.ToString();
            floor.Remove();

            if ((GetCellValue("Calc", "P21") != null) && (double.Parse(GetCellValue("Calc", "P21")) > 0))
            {
                garFlr.Element("Measurements")?.SetAttributeValue("area", Math.Round(garFlrArea * 0.092903, 4));
                garFlr.Element("Measurements")?.SetAttributeValue("length", Math.Round(garFlrLength * 0.3048, 4));
                garFlr.Element("Label")?.SetValue(GetCellValue("Calc", "L21"));
            }
            else
            {
                garFlr.Remove();
            }
        }
        public void ExtraFloors()
        {
            string column = "P";
            int startRow = 22;
            int endRow = 34;
            int currentRow = startRow;
            string area;
            string length;
            string name;

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "O" + currentRow);
                    name = GetCellValue("Calc", "L" + currentRow);

                    NewFloor(name, length, area);
                }
                currentRow++;
            }
        }
        public void ExtraCeilings()
        {
            string column = "E";
            int startRow = 10;
            int endRow = 17;
            int currentRow = startRow;
            string type;
            string length;
            string area;
            string slope;
            string heel;
            string name;
            List<Ceiling> ceilings = new List<Ceiling>();

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "D" + currentRow);
                    name = GetCellValue("Calc", "A" + currentRow);
                    type = GetCellValue("Calc", "C" + currentRow);
                    slope = GetCellValue("Calc", "F" + currentRow);
                    heel = GetCellValue("Calc", "H" + currentRow);
                    ceilings.Add(new Ceiling(name, type, area, length, slope, heel));
                    ceilingCount = ceilings.Count();
                }
                currentRow++;
            }
            ceilingCount = ceilings.Count();
            foreach (Ceiling c in ceilings)
            {
                c.AddCeiling();
            }
        }
        //Checks spreadsheet for vaults and calls AddCeiling() to add them
        public void CheckVaults()
        {
            string column = "M";
            int startRow = 10;
            int endRow = 17;
            int currentRow = startRow;
            string type;
            string length;
            string area;
            string slope;
            string name;
            string rise;
            string heel = GetCellValue("Calc", "H10");
            List<Ceiling> vaults = new List<Ceiling>();

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "L" + currentRow);
                    name = GetCellValue("Calc", "I" + currentRow);
                    type = GetCellValue("Calc", "AD" + currentRow);
                    slope = GetCellValue("Calc", "N" + currentRow);
                    rise = GetCellValue("Calc", "R" + currentRow);
                    vaults.Add(new Ceiling(name, type, area, length, slope, rise, heel, true));
                }
                currentRow++;
            }
            foreach (Ceiling v in vaults)
            {
                v.AddCeiling();
            }
        }

        //Creates new walls that are required but not included in the template. Uses R value from the first floor
        private void NewWall(string name, string corners, string intersections, string height, string perim)
        {
            string heightMetric = Math.Round(Convert.ToDouble(height) * 0.3048, 4).ToString();
            string perimMetric = Math.Round(Convert.ToDouble(perim) * 0.3048, 4).ToString();
            XElement comp = (XElement)(from el in newHouse.Descendants("Components")
                                       select el).First();

            comp.Add(
                new XElement("Wall",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", maxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XAttribute("corners", corners),
                        new XAttribute("intersections", intersections),
                            new XElement("Type", "User specified",
                            new XAttribute("rValue", wallRValue),
                            new XAttribute("nominalInsulation", "3.3527"))),
                    new XElement("Measurements",
                        new XAttribute("height", heightMetric),
                        new XAttribute("perimeter", perimMetric)),
                    new XElement("FacingDirection",
                        new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O"))));
            maxID++;
        }

        //Creates new exposed floors that are not over a garage
        private void NewFloor(string name, string length, string area)
        {
            string lengthMetric = Math.Round(Convert.ToDouble(length) * 0.3048, 4).ToString();
            string areaMetric = Math.Round(Convert.ToDouble(area) * 0.092903, 4).ToString();

            XElement comp = (XElement)(from el in newHouse.Descendants("Components")
                                       select el).First();

            comp.Add(
                new XElement("Floor",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", maxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XElement("Type", "User specified",
                        new XAttribute("rValue", floorRValue),
                        new XAttribute("nominalInsulation", "6.0582"))),
                    new XElement("Measurements",
                        new XAttribute("area", areaMetric),
                        new XAttribute("length", lengthMetric))));
            maxID++;
        }

        //Searches for walls in spreadsheet that aren't included in templates and calls NewWall() to add them
        public void ExtraWalls()
        {
            string column = "H";
            int startRow = 22;
            int endRow = 33;
            int currentRow = startRow;

            for (int i = startRow; i <= endRow; i++)
            {
                string name;
                string corners;
                string intersections;
                string height;
                string perim;
                string currentCell = column + currentRow.ToString();

                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    perim = GetCellValue("Calc", currentCell);
                    height = GetCellValue("Calc", "G" + currentRow);
                    name = GetCellValue("Calc", "A" + currentRow);
                    corners = GetCellValue("Calc", "E" + currentRow);
                    intersections = GetCellValue("Calc", "F" + currentRow);

                    NewWall(name, corners, intersections, height, perim);

                }
                currentRow++;
                if (currentRow == 24)
                {
                    currentRow++;
                    i++;
                }
                if (currentRow == 30)
                {
                    currentRow += 2;
                    i += 2;
                }
            }
        }
        /**
         * 
         * 
         */
        public void ExtractWindows()
        {
            List<Window> windows = new List<Window>();
            for(int i = 2; i <= maxWindowRow; i++)
            {
                string? name = GetCellValue("Windows", "A" + i);
                if (name != null && name != string.Empty && GetCellValue("Windows", "F" + i).ToLower() != "door")
                {
                    int width = int.Parse(GetCellValue("Windows", "B" + i));
                    int height = int.Parse(GetCellValue("Windows", "C" + i));
                    double uValue = double.Parse(GetCellValue("Windows", "D" + i).ToString());
                    double shgc = double.Parse(GetCellValue("Windows", "E" + i));
                    int floor = int.Parse(GetCellValue("Windows", "G" + i));
                    double overhang = double.Parse(GetCellValue("Calc", "M52"));

                    Window window = new Window(name, width, height, uValue, shgc, floor, overhang, maxID);
                    windows.Add(window);
                    window.codeId = CodeTools.FindWindowCode(newHouse, window);
                    window.AddWindow(newHouse);
                    maxID++;
                }
            }
        }

        /**
         * Removes all windows that aren't a part of door assemblies.
         */
        public void RemoveWindows()
        {
            List<XElement>? windows = newHouse?.Root?.Element("House")?.Descendants("Window").
                            Where(el => !el.Ancestors("Door").Any() && el != null).
                            ToList();
            foreach (XElement window in windows)
            {
                window.Remove();
            }
        }
        //Method to get the value of single cells from Excel worksheet
        //Should probably fix this so that the file only opens once
        public static string GetCellValue(string sheetName, string refCell)
        {
            string? value = null;

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart? wbPart = spreadsheet.WorkbookPart;
                    Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().
                        Where(s => s.Name == sheetName).FirstOrDefault();
                    if (theSheet == null)
                    {
                        throw new ArgumentException(null, nameof(sheetName));
                    }
                    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);

                    Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.
                        Where(c => c.CellReference == refCell).FirstOrDefault();
                    if (theCell is null || theCell.InnerText.Length < 0)
                    {
                        return string.Empty;
                    }
                    value = theCell?.CellValue?.InnerText;

                    if (theCell?.DataType != null)
                    {
                        if (theCell.DataType.Value == CellValues.SharedString)
                        {
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (stringTable is not null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                        }
                        else if (theCell.DataType.Value == CellValues.Boolean)
                        {
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                        }
                    }
                }
                fs.Close();
            }return value;
        }
        //stolen regex to separate filename into an address
        static string CamelCaseToSpaceSeparated(string text)
        {
            string[] words = Regex.Matches(text, @"([A-Z]+(?![a-z])|[A-Z][a-z]+|[0-9]+|[a-z]+)")
                .OfType<Match>()
                .Select(m => m.Value)
                .ToArray();
            return string.Join(" ", words);
        }
    }
}
