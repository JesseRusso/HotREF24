//Created by Jesse Russo 2019
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace HotPort
{
    class CreateRef
    {
        private string CeilingRSI { get; set; } = "10.4292";
        private string wallRSI = "3.0802";
        private string garWallRSI = "2.9199";
        private string bsmtWallRSI = "3.3443";
        private string floorRSI = "5.0191";
        private string garFloorRSI = "4.8589";
        private string slabRSI = "1.902";
        private string heatedSlab = "15.8";
        private string windowRSI = "0.62";
        private string doorRSI = "0.6211";
        private string acSeer = "14.5";
        private string furnaceEF = "95";
        private string heatPumpHSPF = "7.1";
        private string defaultACH = "2.5";
        private string defaultUsageBin = "4";
        private int maxID;
        private int codeID = 3;
        private string ExcelFilePath {get; set;}

        public CreateRef(XDocument house, string excelPath, XElement profile)
        {
            List<string> codeIDs = new List<string>();
            var hasCode = from el in house.Descendants("Codes").Descendants().Attributes("id")
                          select el.Value;

            foreach(string code in hasCode)
            {
                string[] codeStrings = code.Split(' ');
                codeIDs.Add(codeStrings[1]);
            }
            codeID = int.Parse(codeIDs.Last()) +1;
            ExcelFilePath = excelPath;
            UpdateValues(profile);
        }
        public void UpdateValues(XElement zoneValues)
        {
            CeilingRSI = zoneValues.Element("CeilingRSI").Value;
            wallRSI = zoneValues.Element("WallRSI").Value;
            garWallRSI = zoneValues.Element("GarWallRSI").Value;
            bsmtWallRSI = zoneValues.Element("BsmtWallRSI").Value;
            floorRSI = zoneValues.Element("FloorRSI").Value;
            slabRSI = zoneValues.Element("SlabRSI").Value;
            heatedSlab = zoneValues.Element("HeatedSlabRSI").Value;
            windowRSI = zoneValues.Element("WindowRSI").Value;
            doorRSI = zoneValues.Element("DoorRSI").Value;
            furnaceEF = zoneValues.Element("FurnaceEF").Value;
            acSeer = zoneValues.Element("ACSeer").Value;
            heatPumpHSPF = zoneValues.Element("HeatPumpHSPF").Value;
            defaultACH = zoneValues.Element("DefaultACH").Value;
            defaultUsageBin = zoneValues.Element("DefaultDrawPattern").Value;
        }

        //Finds the highest value ID attribute in order for new windows and doors to have 
        //their IDs incremented from maxID
        public void FindID(XDocument house)
        {
            List<string> ids = new List<string>();
            var hasID =
                from el in house.Descendants("House").Descendants().Attributes("id")
                where el.Value != null
                select el.Value;
            foreach (string id in hasID)
            {
                ids.Add(id);
            }

            List<int> idList = ids.Select(s => int.Parse(s)).ToList();
            maxID = idList.Max() + 1;
        }
        public XDocument Remover(XDocument house)
        {
            house.Descendants("Door").Remove();
            house.Descendants("Window").Remove();
            house.Descendants("Hrv").Remove();
            house.Descendants("BaseVentilator").Remove();
            return house;
        }
        //Changes R value for ceilings, walls, 
        public XDocument RChanger(XDocument house)
        {
            //Changes R values of ceiling elements
            foreach (XElement ceiling in house.Descendants("Ceiling"))
            {
                string flat = "Flat";
                string cath = "Cathedral";
                foreach (XElement type in ceiling.Descendants("CeilingType"))
                {
                    if (cath.Equals(ceiling.Element("Construction").Element("Type").Element("English").Value.ToString()))
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", floorRSI);
                    }
                    else if (flat.Equals(ceiling.Element("Construction").Element("Type").Element("English").Value.ToString()))
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", floorRSI);
                    } 
                    else
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", CeilingRSI);
                    }
                }
            }
            //Changes R values for walls and rim joists
            foreach (XElement wall in house.Descendants("Wall"))
            {
                //Check for garage walls
                foreach (XElement type in wall.Descendants("Type"))
                    if (wall.Element("Label").Value.ToString().ToLower().Contains("garage"))
                    {
                        //If wall is a garage wall
                        type.SetAttributeValue("rValue", garWallRSI);
                    }
                    else
                    {
                        //If wall is NOT a garage wall
                        type.SetAttributeValue("rValue", wallRSI);
                    }
            }

            //Changes Basement wall insulation and pony wall R values
            foreach (XElement bsmt in house.Descendants("Basement"))
            {
                //Set R value for pony walls if they exist in bsmt
                foreach (XElement pWall in bsmt.Descendants("PonyWallType").Descendants("Section"))
                {
                        pWall.SetAttributeValue("rsi", wallRSI);
                }
                //Set R value for interior wall insulation for bsmt walls
                foreach (XElement intIns in bsmt.Descendants("InteriorAddedInsulation").Descendants("Section"))
                {
                        intIns.SetAttributeValue("rsi", bsmtWallRSI);
                }
                //Set R value for insulation added to slab
                foreach(XElement slab in bsmt.Descendants("AddedToSlab"))
                {
                    slab.SetAttributeValue("rValue", slabRSI);
                }
                //Set R value for bsmt floor headers
                foreach (XElement fh in bsmt.Descendants("FloorHeader").Descendants("Type"))
                {
                        fh.SetAttributeValue("rValue", wallRSI);
                }
            }

            foreach (XElement slab in house.Descendants("Slab").Descendants("Wall").Descendants("RValues"))
            {
                    slab.SetAttributeValue("thermalBreak", slabRSI); 
            }
            foreach (XElement slabIns in house.Descendants("Slab").Descendants("AddedToSlab"))
            {
                slabIns.SetAttributeValue("rValue", slabRSI);
            }

            // Set R value for exposed floors
            string GarFloorName = GetCellValue("Calc", "L21");
            foreach (XElement floor in house.Descendants("Floor"))
            {
                foreach (XElement type in floor.Descendants("Type"))
                {
                    //Check for garage floor
                    if (floor.Element("Label").Value == GarFloorName)
                    {
                        type.SetAttributeValue("rValue", garFloorRSI);
                    }
                    else
                    {
                        type.SetAttributeValue("rValue", floorRSI);
                    }
                }
            }
            return house;
        }
        //Changes values for furnace capacity, furnace EF, ACH, and DHW EF
        public XDocument HeatingCooling(XDocument house)
        {
            //Convert BTUs/h to KW
            double btus = Convert.ToDouble(GetCellValue("General", "C6"));
            btus = Math.Round((btus * 0.00029307107), 5);

            // Changes furnace output capacity and EF values
            foreach (XElement furnace in house.Descendants("Furnace").Descendants("Specifications"))
            {
                    furnace.SetAttributeValue("efficiency", furnaceEF);
                    furnace.SetAttributeValue("isSteadyState", "false");
                    furnace.Element("OutputCapacity").SetAttributeValue("value", btus.ToString());
            }
            // Sets blower fan motor to high efficiency
            XElement fans = house.Descendants("HeatingCooling").Descendants("FansAndPump").FirstOrDefault();
            fans.SetAttributeValue("hasEnergyEfficientMotor", "true");

            //Changes A/C SEER
            foreach (XElement ac in house.Descendants("AirConditioning").Descendants("Efficiency"))
            {
                ac.SetAttributeValue("value", acSeer);
            }
            //Changes heat pump efficiency if HP is present
            foreach (XElement hp in house.Descendants("AirHeatPump"))
            {
                hp.Element("Specifications").Element("HeatingEfficiency").SetAttributeValue("value", heatPumpHSPF);
                hp.Element("Specifications").Element("HeatingEfficiency").SetAttributeValue("isCop", "false");

                if (hp.Element("Equipment").Element("Function").Attribute("code").Value == "2")
                {
                    hp.Element("Specifications").Element("CoolingEfficiency").SetAttributeValue("isCop", "false");
                    hp.Element("Specifications").Element("CoolingEfficiency").SetAttributeValue("value", acSeer);
                }
            }
            return house;
        }
         public XDocument ChangeACH(XDocument house)
        {
            int? houseType = int.Parse(s: house.Element("HouseFile").Element("House").Element("Specifications").Element("HouseType").Attribute("code").Value);
            string? airChanges = house.Element("HouseFile")?.Element("House")?.Element("NaturalAirInfiltration")?.Element("Specifications")?.Element("BlowerTest")?.Attribute("airChangeRate")?.Value.ToString();

            if(houseType != 1 && airChanges != "2.5")
            {
                defaultACH = "3.0";
            }
            //Changes blower door test value to ACH value
            foreach (XElement bt in house.Descendants("BlowerTest"))
            {
                bt.SetAttributeValue("airChangeRate", defaultACH);
            }
            return house;
        }

        public XDocument HotWater(XDocument house)
        {
            foreach (XElement tank in house.Descendants("HotWater").Descendants("Primary"))
            {
                ChangeTank(tank);
            } 
            foreach(XElement tank in house.Descendants("HotWater").Descendants("Secondary"))
            {
                ChangeTank(tank);
            }
            return house;
        }
        //Method to change DHW values
        private void ChangeTank(XElement tank)
        {
            if (tank.Element("EnergySource")?.Attribute("code")?.Value == "2")
            {
                tank.Element("EnergyFactor")?.SetAttributeValue("isUniform", "true");
                tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")), 2));
                Dictionary<string, XElement> map = new()
                {
                    { "1E", new XElement("English", "Very-small-usage 38 L (10 US gal)") },
                    { "1F", new XElement("French", "Très faible utilisation 38 L (10 gal US)") },
                    { "2E", new XElement("English", "Low-usage 144 L (38 US gal)") },
                    { "2F", new XElement("French", "Faible utilisation 144 L (38 gal US)") },
                    { "3E", new XElement("English", "Medium-usage 208 L (55 US gal)") },
                    { "3F", new XElement("French", "Moyenne utilisation 208 L (55 gal US)") },
                    { "4E", new XElement("English", "High-usage 318 L (84 US gal)") },
                    { "4F", new XElement("French", "Grande utilisation 318 L (84 gal US)") },
                };
                if (tank.Element("DrawPattern") == null)
                {
                    tank.Element("EnergyFactor")?.SetAttributeValue("inputCapacity", "0");
                    tank.Element("EnergyFactor").AddAfterSelf(
                        new XElement("DrawPattern",
                            new XAttribute("code", defaultUsageBin),
                            new XElement("English", map[$"{defaultUsageBin}E"].Value),
                            new XElement("French", map[$"{defaultUsageBin}F"].Value)));
                }
                if (tank.Element("TankType").Attribute("code").Value.Equals("12"))
                {
                    tank.SetAttributeValue("flueDiameter", "0");
                    tank.Element("TankType").SetAttributeValue("code", "9");
                    tank.Element("TankVolume").SetAttributeValue("code", "4");
                    tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")), 2));
                }
                return;
            }
            if (tank.Element("EnergySource").Attribute("code").Value == "1")
            {
                tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "J31")), 2));
            }
        }

        //Adds utility ventilation element and fills values for vent rate and fan power
        public XDocument AddFan(XDocument house)
        {
            foreach (XElement vent in house.Descendants("WholeHouseVentilatorList"))
            {
                double ls = Math.Round(Math.Round(Convert.ToDouble(GetCellValue("General", "H4")),1) * 0.47195,4);
                vent.Add(
                    new XElement("BaseVentilator",
                    new XAttribute("supplyFlowrate", ls.ToString()),
                    new XAttribute("exhaustFlowrate", ls.ToString()),
                    new XAttribute("fanPower1", Math.Round(Convert.ToDouble(GetCellValue("General", "K4")),1).ToString()),
                    new XAttribute("isDefaultFanpower", "false"),
                    new XAttribute("isEnergyStar", "false"),
                    new XAttribute("isHomeVentilatingInstituteCertified", "false"),
                    new XAttribute("isSupplemental", "false"),
                        new XElement("EquipmentInformation"),
                        new XElement("VentilatorType", 
                        new XAttribute("code", "4"),
                            new XElement("English", "Utility"),
                            new XElement("French", "Utilité"))));
            }
            return house;
        }

        public XDocument Doors(XDocument house)
        {
            double width = Math.Round((Convert.ToDouble(GetCellValue("General", "N10")) * 0.0254), 4);
            string ff = "1st Flr";
            foreach (XElement wall in house.Descendants("Wall"))
            {
                foreach (XElement comp in wall.Descendants("Components"))
                {
                    if (ff.Equals(wall.Element("Label")?.Value))
                    {
                        int i = 0;
                        while (i < 2)
                        {
                            comp.Add(
                                new XElement("Door",
                                new XAttribute("rValue", doorRSI),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "Door"),
                                    new XElement("Construction",
                                    new XAttribute("energyStar", "false"),
                                        new XElement("Type",
                                        new XAttribute("code", "8"),
                                        new XAttribute("value", doorRSI),
                                            new XElement("English", "User Specified"),
                                            new XElement("French", "Spécifié par l'utilisateur"))),
                                    new XElement("Measurements",
                                    new XAttribute("height", "2.1336"),
                                    new XAttribute("width", width))));
                            i++;
                            maxID++;
                        }
                    }
                }
            }
            return house;
        }

        public XDocument Windows(XDocument house)
        {
            double size = Math.Round((System.Convert.ToDouble(GetCellValue("General", "N9")) * 25.4), 6);
            string floors = "2nd Flr";
            List<string> wallList = new List<string>();
            Dictionary<string, string> facingDirection = new()
            {
                { "North", "5" },
                { "South", "1" },
                { "West", "7" },
                { "East", "3" },
            };

            var walls =
                from el in house.Descendants("House").Descendants("Wall").Descendants("Label")
                where el.Value != null
                select el.Value.ToString();

            foreach (string wall in walls)
            {
                wallList.Add(wall);
            }

            //Checks if a second floor exists. If not, windows are added to the first floor
            if (!wallList.Contains("2nd Flr"))
            {
                floors = "1st Flr";
            }
            foreach (XElement wall in house.Descendants("Wall"))
            {
                foreach (XElement comp in wall.Descendants("Components"))
                {
                    if (floors.Equals(wall?.Element("Label")?.Value))
                    {
                        foreach (KeyValuePair<string, string> pair in facingDirection)
                        {
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", pair.Key),
                                    new XElement("Construction",
                                    new XAttribute("energyStar", "false"),
                                        new XElement("Type", "ABCRef",
                                        new XAttribute("idref", ("Code " + codeID)),
                                        new XAttribute("rValue", windowRSI))),
                                    new XElement("Measurements",
                                    new XAttribute("height", size),
                                    new XAttribute("width", size),
                                    new XAttribute("headerHeight", "0"),
                                    new XAttribute("overhangWidth", "0"),
                                        new XElement("Tilt",
                                        new XAttribute("code", "1"),
                                        new XAttribute("value", "90"),
                                        new XElement("English", "Vertical"),
                                        new XElement("French", "Verticale"))),
                                    new XElement("Shading",
                                    new XAttribute("curtain", "1"),
                                    new XAttribute("shutterRValue", "0")),
                                    new XElement("FacingDirection",
                                    new XAttribute("code", pair.Value),
                                    new XElement("English", "North"),
                                    new XElement("French", "Nord"))));
                            maxID++;
                        }
                    }
                }
            }
            return house;
        }
        public XDocument AddCode(XDocument house)
            {
                var codes = from el in house.Descendants()
                            where el.Name == "Codes"
                            select el;

                foreach (XElement code in codes)
                {
                    code.AddFirst(
                        new XElement("Window",
                            new XElement("UserDefined",
                                new XElement("Code",
                                new XAttribute("id", ("Code " + codeID)),
                                new XAttribute("nominalRValue", "0"),
                                    new XElement("Label", "ABCRef"),
                                    new XElement("Description", ""),
                                    new XElement("Layers",
                                        new XElement("WindowLegacy",
                                        new XAttribute("frameHeight", "0"),
                                        new XAttribute("shgc", "0.26"),
                                        new XAttribute("rank", "1"),
                                            new XElement("Type",
                                            new XAttribute("code", 1),
                                                new XElement("English", "Picture"),
                                                new XElement("French", "Fixe")),
                                            new XElement("RsiValues",
                                            new XAttribute("centreOfGlass", windowRSI),
                                            new XAttribute("edgeOfGlass", windowRSI),
                                            new XAttribute("frame", windowRSI))))))));
                }
            return house;
        }
        //Method to get the value of a single cell from worksheet
        public string GetCellValue(string sheetName, string refCell)
        {
            string? value = null;

            using (FileStream fs = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart? wbPart = spreadsheet.WorkbookPart;
                    Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().
                        Where(s => s.Name == sheetName).FirstOrDefault();
                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
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
            }
            return value;
        }
    }
}