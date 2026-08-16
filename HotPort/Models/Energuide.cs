using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using HotPort.Properties;

namespace HotPort.Models
{
    public class Energuide
    {
        public XDocument House { get; }

        private readonly string _excelFilePath;
        public string Builder { get; private set; } = string.Empty;
        public int MaxID { get; private set; }
        private bool _basementPresent = true;

        private readonly CodeEntry? _mainWallCode;
        private readonly CodeEntry? _noCladWallCode;
        private readonly CodeEntry? _tallWallCode;
        private readonly CodeEntry? _floorHeaderCode;
        private readonly CodeEntry? _ceilingCode;
        private readonly CodeEntry? _vaultCode;
        private readonly CodeEntry? _cathedralCode;
        private readonly CodeEntry? _flatCode;
        private readonly CodeEntry? _exposedFloorCode;
        private readonly CodeEntry? _garageFloorCode;
        private readonly CodeEntry? _floorsAboveCode;
        private readonly CodeEntry? _interiorWallCode;
        private readonly CodeEntry? _ponyWallCode;

        private readonly Dictionary<string, string> _assignedIds = new();
        private readonly List<CodeEntry> _usedCodes = new();
        private int _nextCodeId = 1;

        public Energuide(
            XDocument template,
            string excelFilePath,
            CodeEntry? mainWallCode,
            CodeEntry? noCladWallCode,
            CodeEntry? tallWallCode,
            CodeEntry? floorHeaderCode,
            CodeEntry? ceilingCode,
            CodeEntry? vaultCode,
            CodeEntry? cathedralCode,
            CodeEntry? flatCode,
            CodeEntry? exposedFloorCode,
            CodeEntry? garageFloorCode,
            CodeEntry? floorsAboveCode,
            CodeEntry? interiorWallCode,
            CodeEntry? ponyWallCode)
        {
            House = new XDocument(template);
            _excelFilePath = excelFilePath;
            _mainWallCode = mainWallCode;
            _noCladWallCode = noCladWallCode;
            _tallWallCode = tallWallCode;
            _floorHeaderCode = floorHeaderCode;
            _ceilingCode = ceilingCode;
            _vaultCode = vaultCode;
            _cathedralCode = cathedralCode;
            _flatCode = flatCode;
            _exposedFloorCode = exposedFloorCode;
            _garageFloorCode = garageFloorCode;
            _floorsAboveCode = floorsAboveCode;
            _interiorWallCode = interiorWallCode;
            _ponyWallCode = ponyWallCode;
        }

        // ── Entry points ───────────────────────────────────────────────────────

        public void Generate()
        {
            BuildHouse();
            BuildCodesSection();
            if (Settings.Default.ErsWindowsCheckbox)
            {
                RemoveWindows();
                ExtractWindows();
            }
        }

        public void BuildHouse()
        {
            GetBuilder();
            FindComponentID();
            _nextCodeId = CodeTools.GetValidCodeID(House);
            CityCheck();
            SetDate();
            SetFileId();
            ChangeCityWeather();
            ChangeSpecs();
            ChangeFurnace();
            CheckAC();
            ProcessWalls();
            SetHouseType();
            SetStoreys();
            RemoveCeilings();
            ExtraCeilings();
            CheckVaults();
            CheckFloors();
            CheckBasement();
            ChangeHRV();
            GasDHW();
            ElectricDHW();
            CheckFireplace();
            AddDoors();
            SolarPanels();
        }

        public void ChangeAddress(string address)
        {
            foreach (XElement add in House.Descendants("Client").Descendants("StreetAddress"))
            {
                add.Element("Street")?.SetValue(address);
                string postCode = GetCellValue("Calc", "P2");
                if (!string.IsNullOrEmpty(postCode))
                    add.SetElementValue("PostalCode", postCode);
            }
            XElement addressEl = House.Element("HouseFile").Element("ProgramInformation").Element("Client").Element("StreetAddress").Element("Street");
            addressEl.SetValue(GetCellValue("Calc", "P5"));
        }

        // ── Infrastructure ─────────────────────────────────────────────────────

        private void GetBuilder()
        {
            Builder = House.Descendants("File").Descendants("BuilderName")
                          .Where(el => !el.IsEmpty)
                          .First().Value;
        }

        private void FindComponentID()
        {
            var ids = House.Descendants("House").Descendants().Attributes("id")
                          .Where(a => a.Value != null)
                          .Select(a => int.Parse(a.Value))
                          .ToList();
            MaxID = ids.Max() + 1;
        }
        private int AssignComponentID()
        {
            int id = MaxID;
            MaxID++;
            return id;
        }

        private string GetCellValue(string sheet, string cell) =>
            ExcelHelper.GetCellValue(_excelFilePath, sheet, cell);

        private double GetDoubleCellValue(string sheet, string cell) =>
            ExcelHelper.GetDoubleCellValue(_excelFilePath, sheet, cell);

        // ── Location & specs ───────────────────────────────────────────────────

        private void CityCheck()
        {
            string city = GetCellValue("Calc", "P1");
            if (string.IsNullOrEmpty(city))
                throw new FormatException("!Calc P1 cannot be empty. Please select a city.");


            //Changes the Terrain dropdown in the NAI screen to the Okotoks value
            if (city == "Okotoks")
                House.Element("HouseFile").Element("House").Element("NaturalAirInfiltration")
                     .Element("Specifications").Element("BuildingSite").Element("Terrain")
                     .SetAttributeValue("code", 3);

            House.Element("HouseFile").Element("ProgramInformation").Element("Client")
                 .Element("StreetAddress").Element("City").SetValue(city);
        }

        private void ChangeCityWeather()
        {
            string weatherStation = GetCellValue("Calc", "P6");
            if (string.IsNullOrEmpty(weatherStation))
                throw new ArgumentException("Invalid data found in the Weather zone field; !Calc P6");

            int locationCode = weatherStation.ToLower() switch
            {
                "cop upper" => 10,
                "stavely"   => 37,
                "sundre"    => 38,
                _           => 5
            };

            House.Element("HouseFile").Element("ProgramInformation").Element("Weather")
                 .Element("Location").SetAttributeValue("code", locationCode);
        }
        /*
         * Changes the above/below ground heated floor area in the H2K specifications screen, 
         * and the house volume and height of highest ceiling in the Natural Air Infiltration screen.
         */
        private void ChangeSpecs()
        {
            double aboveGrade = GetDoubleCellValue("Calc", "E6");
            double belowGrade = GetDoubleCellValue("Calc", "F6");

            foreach (XElement grade in House.Descendants("Specifications").Descendants("HeatedFloorArea"))
            {
                grade.SetAttributeValue("aboveGrade", Math.Round(aboveGrade * 0.092903, 1));
                grade.SetAttributeValue("belowGrade", Math.Round(belowGrade * 0.092903, 1));
            }

            foreach (XElement infil in House.Descendants("NaturalAirInfiltration"))
            {
                infil.Element("Specifications").Element("House")
                     .SetAttributeValue("volume", Math.Round(GetDoubleCellValue("Calc", "Q49") * 0.02831684609, 4).ToString());
                infil.Element("Specifications").Element("BuildingSite")
                     .SetAttributeValue("highestCeiling", Math.Round(GetDoubleCellValue("Calc", "M51") * 0.3048, 4).ToString());
            }

            var corners = new List<int>();
            for (int i = 2; i <= 5; i++)
            {
                if (int.TryParse(GetCellValue("Calc", $"F{i}"), out int count))
                    corners.Add(count);
            }
            int maxCorners = corners.Any() ? corners.Max() : 4;

            foreach (XElement ps in House.Descendants("PlanShape"))
            {
                (int code, string eng, string fr) = maxCorners switch
                {
                    <= 4  => (1, "Rectangular", "Rectangulaire"),
                    <= 6  => (4, "Other, 5-6 corners", "Autre, 5-6 coins"),
                    <= 8  => (5, "Other, 7-8 corners", "Autre, 7-8 coins"),
                    <= 10 => (6, "Other, 9-10 corners", "Autre, 9-10 coins"),
                    _     => (7, "Other, 11 or more corners", "Autre, 11 coins ou plus")
                };
                ps.SetAttributeValue("code", code.ToString());
                ps.Element("English")?.SetValue(eng);
                ps.Element("French")?.SetValue(fr);
            }
        }

        private void SetDate()
        {
            XElement evalDate = House.Element("HouseFile").Element("ProgramInformation").Element("File");
            evalDate.SetAttributeValue("evaluationDate", DateTime.UtcNow.Date.ToString("yyyy-MM-dd"));
        }

        private void SetFileId()
        {
            string fileId = GetCellValue("Calc", "P3");
            XElement idField = House.Element("HouseFile").Element("ProgramInformation").Element("File").Element("Identification");
            if (fileId != null)
                idField.SetValue(fileId);
        }

        private void SetStoreys()
        {
            bool firstFloorHt  = double.TryParse(GetCellValue("Calc", "E3"), out double fh) && fh > 0;
            bool firstPerim    = double.TryParse(GetCellValue("Calc", "C3"), out double fp) && fp > 0;
            bool secondFloorHt = double.TryParse(GetCellValue("Calc", "E4"), out double sh) && sh > 0;
            bool secondPerim   = double.TryParse(GetCellValue("Calc", "C4"), out double sp) && sp > 0;
            bool thirdFloorHt  = double.TryParse(GetCellValue("Calc", "E5"), out double th) && th > 0;
            bool thirdPerim    = double.TryParse(GetCellValue("Calc", "C5"), out double tp) && tp > 0;

            XElement spec = House.Root.Element("House").Element("Specifications");
            if (firstFloorHt && firstPerim)
            {
                if (secondFloorHt && secondPerim)
                    spec.Element("Storeys").SetAttributeValue("code", thirdFloorHt && thirdPerim ? "5" : "3");
                else
                    spec.Element("Storeys").SetAttributeValue("code", "1");
            }
        }

        private void SetHouseType()
        {
            foreach (XElement houseType in House.Descendants("HouseType"))
            {
                string? calcK5 = GetCellValue("Calc", "K5")?.ToLower();
                if (calcK5 == "single") continue;

                switch (calcK5)
                {
                    case "semi":
                        houseType.SetAttributeValue("code", "2");
                        break;
                    case "rowhouse-end":
                        houseType.SetAttributeValue("code", "6");
                        break;
                    case "rowhouse-mid":
                        houseType.SetAttributeValue("code", "8");
                        break;
                    case "murb duplex detached":
                        houseType.SetAttributeValue("code", "9");
                        AddMurbData();
                        break;
                    case "murb duplex attached":
                        houseType.SetAttributeValue("code", "11");
                        AddMurbData();
                        break;
                    default:
                        houseType.SetAttributeValue("code", "1");
                        break;
                }
            }
        }

        private void AddMurbData()
        {
            XElement? house        = House.Root.Element("House");
            XElement? specifications = house?.Element("Specifications");
            XElement? infilFactors = house?.Element("NaturalAirInfiltration").Element("OtherFactors");

            house?.Element("Labels").SetElementValue("English", "Multi-unit: one unit");
            specifications?.SetAttributeValue("buildingType", "Multi-unit: one unit");

            specifications?.Element("HeatedFloorArea").AddBeforeSelf(
                new XElement("NumberOf",
                    new XAttribute("storeysInBuilding", "2"),
                    new XAttribute("dwellingUnits", "2"),
                    new XAttribute("nonResUnits", "0")));

            infilFactors?.Element("LeakageFractions")?.SetAttributeValue("ceilings", "0.3");
            infilFactors?.Element("LeakageFractions")?.SetAttributeValue("walls", "0.5");
            infilFactors?.Element("LeakageFractions")?.SetAttributeValue("floors", "0.2");
        }

        // ── Equipment ──────────────────────────────────────────────────────────

        private void ChangeHRV()
        {
            string hrvMake   = GetCellValue("Summary", "D74");
            string hrvModel  = GetCellValue("Summary", "E74");
            string hrvPower1 = GetCellValue("General", "I4");
            string hrvPower2 = GetCellValue("General", "J4");
            string hrvSRE1   = GetCellValue("General", "I5");
            string hrvSRE2   = GetCellValue("General", "J5");

            if (int.TryParse(GetCellValue("General", "K2"), out int count) && count > 0)
                hrvModel += " & " + GetCellValue("Summary", "E92");
            if (int.TryParse(GetCellValue("General", "K1"), out int qty) && qty > 1)
                hrvModel = $"{qty}X {hrvModel}";
            double hrvFlowrate = Math.Round(GetDoubleCellValue("General", "H4"), 1);
            string fanPower    = Math.Round(GetDoubleCellValue("General", "K4"), 1).ToString();

            if (GetDoubleCellValue("General", "I4") <= 0)
            {
                foreach (XElement vent in House.Descendants("WholeHouseVentilatorList"))
                {
                    vent.Element("Hrv")?.Remove();
                    vent.Add(HRV.CreateFan(hrvFlowrate, fanPower));
                }
            }
            else
            {
                foreach (XElement hrv in House.Descendants("WholeHouseVentilatorList"))
                {
                    hrv.Element("Hrv").Element("EquipmentInformation").Element("Manufacturer").SetValue(hrvMake);
                    hrv.Element("Hrv").Element("EquipmentInformation").Element("Model").SetValue(hrvModel);
                    hrv.Element("Hrv").SetAttributeValue("supplyFlowrate", Math.Round(hrvFlowrate * 0.471947, 4).ToString());
                    hrv.Element("Hrv").SetAttributeValue("exhaustFlowrate", Math.Round(hrvFlowrate * 0.471947, 4).ToString());
                    hrv.Element("Hrv").SetAttributeValue("fanPower1", hrvPower1);
                    hrv.Element("Hrv").SetAttributeValue("fanPower2", hrvPower2);
                    hrv.Element("Hrv").SetAttributeValue("efficiency1", hrvSRE1);
                    hrv.Element("Hrv").SetAttributeValue("efficiency2", hrvSRE2);

                    if (_basementPresent)
                    {
                        hrv.Element("Hrv").Element("ColdAirDucts").Element("Supply").Element("Location").SetAttributeValue("code", "1");
                        hrv.Element("Hrv").Element("ColdAirDucts").Element("Exhaust").Element("Location").SetAttributeValue("code", "1");
                    }
                    else
                    {
                        hrv.Element("Hrv").Element("ColdAirDucts").Element("Supply").Element("Location").SetAttributeValue("code", "4");
                        hrv.Element("Hrv").Element("ColdAirDucts").Element("Exhaust").Element("Location").SetAttributeValue("code", "4");
                    }
                }
            }
        }

        private void ChangeFurnace()
        {
            string furnaceModel = GetCellValue("Summary", "B78");
            double furnaceBtus  = GetDoubleCellValue("General", "C4");

            if (GetDoubleCellValue("General", "B4") > 0)
                furnaceModel += $" & {GetCellValue("Summary", "B79")}";

            //Sets the values in the Furnace screen
            foreach (XElement furn in House.Descendants("Furnace"))
            {
                furn.Element("Specifications")?.SetAttributeValue("efficiency", GetCellValue("General", "A6"));
                furn.Element("Specifications")?.Element("OutputCapacity")
                    ?.SetAttributeValue("value", Math.Round(furnaceBtus * 0.00029307107, 5).ToString());
                furn.Element("EquipmentInformation").Element("Manufacturer").SetValue(GetCellValue("Summary", "A78"));
                furn.Element("EquipmentInformation").Element("Model").SetValue(furnaceModel);

                if (Builder.ToLower().Contains("trico"))
                {
                    furn.Element("Specifications").Element("OutputCapacity").SetAttributeValue("code", "2");
                }
            }
            //Sets the values in the "Fans/Pumps" screen
            foreach (XElement fan in House.Descendants("HeatingCooling"))
            {
                fan.Element("Type1").Element("FansAndPump").SetAttributeValue("hasEnergyEfficientMotor", "true");
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("isCalculated", "true");
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("low", "0");
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("high", "300");

                fan.Element("Type1").Element("FansAndPump").Element("Mode").SetAttributeValue("code", "1");
            }
        }

        private void CheckAC()
        {
            string ACSeer = GetCellValue("General", "O26");
            if (!double.TryParse(ACSeer, out double seer) || seer <= 0) return;

            string make = GetCellValue("General", "N25");
            string model = GetCellValue("General", "P25");
            string ahri = GetCellValue("General", "R26");
            string indoorCoil = GetCellValue("General", "T26");
            string seer2 = GetCellValue("General", "Q26").ToLower();

            if(ahri != null || ahri != string.Empty)
            {
                make += $" AHRI#{ahri}";
            }
            if(indoorCoil != null || indoorCoil != string.Empty)
            {
                model += $" / {indoorCoil}";
            }

            House?.Element("HouseFile")?.Element("House")?.Element("Temperatures")?.Element("Basement")
                 ?.SetAttributeValue("cooled", "true");

            XElement? type2 = House?.Descendants("HeatingCooling").Descendants("Type2").FirstOrDefault();
            XElement airCon =
                new XElement("AirConditioning",
                    new XElement("EquipmentInformation",
                        new XAttribute("energystar", "false"),
                        new XElement("Manufacturer", make),
                        new XElement("Model", model)),
                    new XElement("Equipment",
                        new XAttribute("crankcaseHeater", "60"),
                        new XElement("CentralType", new XAttribute("code", "1"))),
                    new XElement("Specifications",
                        new XAttribute("sizingFactor", "1"),
                        new XElement("RatedCapacity",
                            new XAttribute("code", "2"),
                            new XAttribute("value", "2.015"),
                            new XAttribute("uiUnits", "btu/hr")),
                        new XElement("Efficiency",
                            new XAttribute("isCop", "false"),
                            new XAttribute("value", seer))),
                    new XElement("CoolingParameters",
                        new XAttribute("sensibleHeatRatio", "0.76"),
                        new XAttribute("openableWindowArea", "0"),
                        new XElement("FansAndPump",
                            new XAttribute("flowRate", "123"),
                            new XAttribute("hasEnergyEfficientMotor", "false"),
                            new XElement("Mode", new XAttribute("code", "1")),
                            new XElement("Power", new XAttribute("isCalculated", "true")))));
            if(seer2 == "y")
            {
                airCon.Element("Specifications").Element("Efficiency").SetAttributeValue("unit", "2");
            }
            type2.Add(airCon);
        }

        // ── Walls ──────────────────────────────────────────────────────────────

        private void ProcessWalls()
        {
            // Remove any placeholder wall the template may contain
            House.Descendants("Components").Elements("Wall").ToList().ForEach(w => w.Remove());

            for (int row = 21; row <= 34; row++)
            {
                string name = GetCellValue("Calc", $"A{row}");
                if (string.IsNullOrEmpty(name)) continue;

                double perimeter = GetDoubleCellValue("Calc", $"H{row}");
                if (perimeter <= 0) continue;

                string corners       = GetCellValue("Calc", $"E{row}");
                string intersections = GetCellValue("Calc", $"F{row}");
                double height        = GetDoubleCellValue("Calc", $"G{row}");
                string fhPerim       = GetCellValue("Calc", $"J{row}");
                string fhHeight      = GetCellValue("Calc", $"K{row}");
                string spacing       = GetCellValue("Calc", $"B{row}") ?? string.Empty;
                string siding        = GetCellValue("Calc", $"D{row}") ?? string.Empty;
                bool uncon           = GetCellValue("Calc", $"I{row}").ToLower() == "y";

                AddNewWall(name, corners, intersections, height, perimeter, uncon);

                XElement? newWall = House.Descendants("Wall")
                    .FirstOrDefault(w => w.Attribute("id")?.Value == (MaxID - 1).ToString());
                XElement? wallTypeEl = newWall?.Element("Construction")?.Element("Type");
                CodeEntry? wallCode = WallCodeFromSpecs(spacing, siding);

                if (wallTypeEl != null && wallCode != null)
                    ApplyCode(wallTypeEl, wallCode, includeNominalInsulation: true);

                if (Double.TryParse(fhPerim, out double headerPerim) && headerPerim > 0 && Double.TryParse(fhHeight, out double headerHeight) && headerHeight > 0)
                {
                    if (newWall?.Element("Components") == null)
                        newWall!.Add(new XElement("Components"));
                    newWall?.Element("Components")?.Add(
                        FloorHeader.NewJoist(fhHeight.ToString(), "0", fhPerim.ToString(), AssignComponentID().ToString()));

                    if (_floorHeaderCode != null)
                    {
                        XElement? fhTypeEl = newWall?.Element("Components")?.Element("FloorHeader")
                            ?.Element("Construction")?.Element("Type");
                        if (fhTypeEl != null)
                            ApplyCode(fhTypeEl, _floorHeaderCode, includeNominalInsulation: true);
                    }
                }
            }
        }

        private CodeEntry? WallCodeFromSpecs(string spacing, string siding) =>
            siding.ToLower().Contains("n/a") ? _noCladWallCode :
            spacing.Contains("12") ? _tallWallCode :_mainWallCode;

        private void AddNewWall(string name, string corners, string intersections, double height, double perim, bool adjacentUnconditioned)
        {
            House.Descendants("Components").First().Add(
                new XElement("Wall",
                    new XAttribute("adjacentEnclosedSpace", adjacentUnconditioned.ToString().ToLower()),
                    new XAttribute("id", AssignComponentID()),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XAttribute("corners", corners),
                        new XAttribute("intersections", intersections),
                        new XElement("Type", "",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("height",    Math.Round(height * 0.3048, 4).ToString()),
                        new XAttribute("perimeter", Math.Round(perim  * 0.3048, 4).ToString())),
                    new XElement("FacingDirection",
                        new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O"))));
        }

        // ── Ceilings ───────────────────────────────────────────────────────────

        private void RemoveCeilings()
        {
            // Remove all template ceiling placeholders; actual ceilings are added by ExtraCeilings/CheckVaults
            House.Descendants("Ceiling")
                 .ToList().ForEach(c => c.Remove());
        }

        private void ExtraCeilings()
        {
            for (int row = 10; row <= 16; row++)
            {
                string cellValue = GetCellValue("Calc", $"E{row}");
                if (!double.TryParse(cellValue, out double area) || area <= 0) continue;

                double length = GetDoubleCellValue("Calc", $"D{row}");
                string name   = GetCellValue("Calc", $"A{row}");
                string type   = GetCellValue("Calc", $"C{row}");
                string slope  = GetCellValue("Calc", $"F{row}");
                double heel   = GetDoubleCellValue("Calc", $"H{row}");

                AddCeiling(name, type, area, length, slope, heel, isVault: false);
            }
        }

        private void CheckVaults()
        {
            double heel = GetDoubleCellValue("Calc", "H10");

            // Collect every vault row from the spreadsheet
            var vaults = new List<(string name, string type, double area, double length, string slope)>();
            for (int row = 10; row <= 17; row++)
            {
                string cellValue = GetCellValue("Calc", $"M{row}");
                if (string.IsNullOrEmpty(cellValue) || !double.TryParse(cellValue, out double area) || area <= 0)
                    continue;

                double length = GetDoubleCellValue("Calc", $"L{row}");
                string name   = GetCellValue("Calc", $"I{row}");
                string type   = GetCellValue("Calc", $"AD{row}");
                string slope  = GetCellValue("Calc", $"N{row}");

                vaults.Add((name, type, area, length, slope));
            }

            // Merge vaults sharing the same slope into a single vault (sum area + length)
            foreach (var group in vaults.GroupBy(v => v.slope))
            {
                var first = group.First();
                double totalArea   = group.Sum(v => v.area);
                double totalLength = group.Sum(v => v.length);
                string mergedName  = "Vault " + string.Join("/", group.Select(v => v.name.Split(" ").Last()).Distinct());

                AddCeiling(mergedName, first.type, totalArea, totalLength, group.Key, heel, isVault: true);
            }
        }

        private void AddCeiling(string name, string type, double area, double length, string slope, double heel, bool isVault)
        {
            (string typeCode, string typeEng, string typeFr) = (type?.ToLower()) switch
            {
                "gable"     => ("2", "Attic/gable",  "Combles/pignon"),
                "hip"       => ("3", "Attic/hip",    "Combles/arête"),
                "cathedral" => ("4", "Cathedral",    "Cathédrale"),
                "flat"      => ("5", "Flat",         "Plat"),
                "scissor"   => ("6", "Scissor",      "Ciseaux"),
                _           => ("3", "Attic/hip",    "Combles/arête")
            };

            bool isFlat = type?.ToLower() == "flat";
            (string slopeCode, string slopeValue, string slopeEng, string slopeFr) = ResolveSlope(slope, isFlat);

            string heelHeight  = Math.Round(heel   * 0.3048,   3).ToString();
            string areaMetric  = Math.Round(area   * 0.092903, 4).ToString();
            string lengthMetric = Math.Round(length * 0.3048,  4).ToString();

            string slopeName = isFlat || slope == "0" ? "Flat"
                             : isVault ? string.Empty
                             : $"{slope}/12";

            XElement ceiling =
                new XElement("Ceiling",
                    new XAttribute("id", AssignComponentID()),
                    new XElement("Label", $"{name} {slopeName}".TrimEnd()),
                    new XElement("Construction",
                        new XElement("Type",
                            new XAttribute("code", typeCode),
                            new XElement("English", typeEng),
                            new XElement("French",  typeFr)),
                        new XElement("CeilingType", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("length",     lengthMetric),
                        new XAttribute("area",       areaMetric),
                        new XAttribute("heelHeight", heelHeight),
                        new XElement("Slope",
                            new XAttribute("code",  slopeCode),
                            new XAttribute("value", slopeValue),
                            new XElement("English", slopeEng),
                            new XElement("French",  slopeFr))));

            House.Descendants("Components").First().Add(ceiling);

            CodeEntry? code = typeCode switch
            {
                "2" or "3" => _ceilingCode,
                "6"        => _vaultCode,
                "4"        => _cathedralCode,
                "5"        => _flatCode,
                _          => null
            };
            if (code != null)
                ApplyCode(ceiling.Element("Construction").Element("CeilingType"), code, includeNominalInsulation: true);
        }

        private static (string code, string value, string eng, string fr) ResolveSlope(string slope, bool isFlat)
        {
            if (isFlat || slope == "0")
                return ("1", "0", "Flat roof", "Toit plat");

            if (double.TryParse(slope, out double slopeNum) && slopeNum > 7)
                return ("0", Math.Round(slopeNum / 12, 4).ToString(), "User specified", "Spécifié par l'utilisateur");

            return slope switch
            {
                "2" => ("2", "0.167", "2 / 12", "2 / 12"),
                "3" => ("3", "0.25",  "3 / 12", "3 / 12"),
                "4" => ("4", "0.333", "4 / 12", "4 / 12"),
                "5" => ("5", "0.417", "5 / 12", "5 / 12"),
                "6" => ("6", "0.5",   "6 / 12", "6 / 12"),
                "7" => ("7", "0.583", "7 / 12", "7 / 12"),
                _   => ("0", Math.Round(Convert.ToDouble(slope) / 12, 4).ToString(), "User specified", "Spécifié par l'utilisateur")
            };
        }

        // ── Floors ─────────────────────────────────────────────────────────────

        private void CheckFloors()
        {

            IEnumerable<XElement> flrs = House.Descendants("Components").Elements("Floor");

            foreach (XElement flr in flrs)
            {
                flr.Remove();
            }
            for (int row = 21; row <= 30; row++)
            {
                string cellValue = GetCellValue("Calc", $"P{row}");
                if (string.IsNullOrEmpty(cellValue) || !double.TryParse(cellValue, out double area) || area <= 0)
                    continue;

                double length = GetDoubleCellValue("Calc", $"O{row}");
                string name   = GetCellValue("Calc", $"L{row}");
                bool gar = GetCellValue("Calc", $"Q{row}").ToLower() == "y";
                NewFloor(name, length, area, gar);
            }
        }

        private void NewFloor(string name, double length, double area, bool gar)
        {
            XElement floor =
                new XElement("Floor",
                    new XAttribute("adjacentEnclosedSpace", gar ? "true" : "false"),
                    new XAttribute("id", AssignComponentID()),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XElement("Type", "User specified",
                            new XAttribute("idref", ""),
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("area",   Math.Round(area   * 0.092903, 4).ToString()),
                        new XAttribute("length", Math.Round(length * 0.3048,   4).ToString())));

            House.Descendants("Components").First().Add(floor);

            CodeEntry? code = gar ? _garageFloorCode : _exposedFloorCode;
            if (code != null)
                ApplyCode(floor.Element("Construction").Element("Type"), code, includeNominalInsulation: true);
        }

        // ── Basement ───────────────────────────────────────────────────────────

        public void CheckBasement()
        {
            bool basement = GetCellValue("Calc", "N2")?.ToLower() == "y";

            XElement? existingBsmt = House.Descendants("Components").Descendants("Basement").FirstOrDefault();
            if (existingBsmt != null)
            {
                existingBsmt.Remove();
            }

            if (basement)
            {
                _basementPresent = true;

                double pony = Math.Round(GetDoubleCellValue("Calc", "I38") * 0.3048, 4);
                string corners = GetCellValue("Calc", "J38");
                double expPerim = Math.Round(Convert.ToDouble(GetCellValue("Calc", "E38")) * 0.3048, 4);
                double floorPerim = Math.Round(GetDoubleCellValue("Calc", "D38") * 0.3048, 4);
                double floorArea = Math.Round(GetDoubleCellValue("Calc", "F38") * 0.092903, 4);
                double wallHeight = Math.Round(GetDoubleCellValue("Calc", "G38") * 0.3048, 4);
                double depthBelow = Math.Round(GetDoubleCellValue("Calc", "H38") * 0.3048, 4);
                double headerHeight = Math.Round(GetDoubleCellValue("Calc", "K38") * 0.3048, 4);
                double headerPerim = Math.Round(GetDoubleCellValue("Calc", "L38") * 0.3048, 4);

                Foundation fnd = new Foundation
                {
                    Id = AssignComponentID(),
                    Label = "Basement",
                    ExposedSurfacePerimeter = expPerim,
                    FloorPerimeter = floorPerim,
                    FloorArea = floorArea,
                    WallCorners = corners,
                    WallHeight = wallHeight,
                    WallDepth = depthBelow,
                    PonyWallHeight = pony,
                };
                XElement bsmt = fnd.AddBasement();

                // Floors above: a leaf <FloorsAbove> element, coded like any other Type
                XElement? floorsAbove = bsmt.Element("Floor")?.Element("Construction")?.Element("FloorsAbove");
                if (floorsAbove != null && _floorsAboveCode != null)
                    ApplyCode(floorsAbove, _floorsAboveCode, includeNominalInsulation: true);
                if (_floorsAboveCode.NominalRValue <= 0)
                    floorsAbove.SetAttributeValue("rValue", "0.4");

                // Interior wall insulation has its own structure (Description + Composite), so set it directly
                XElement? interior = bsmt.Element("Wall")?.Element("Construction")?.Element("InteriorAddedInsulation");
                if (interior != null && _interiorWallCode != null)
                {
                    string rValue = _interiorWallCode.NominalRValue.ToString();
                    interior.SetAttributeValue("idref", GetOrAssignId(_interiorWallCode));
                    interior.SetAttributeValue("nominalInsulation", rValue);
                    interior.Element("Description")?.SetValue(_interiorWallCode.Label);
                    interior.Element("Composite")?.Element("Section")?.SetAttributeValue("nominalRsi", rValue);
                }
                // Adds a pony wall to the basemnt if one is present
                if(pony > 0D)
                {
                    fnd.AddPonyWall(bsmt);

                    // PonyWallType is the same structure as InteriorAddedInsulation, so code it the same way
                    XElement? ponyType = bsmt.Element("Wall")?.Element("Construction")?.Element("PonyWallType");
                    if (ponyType != null && _ponyWallCode != null)
                    {
                        string rValue = _ponyWallCode.NominalRValue.ToString();
                        ponyType.SetAttributeValue("idref", GetOrAssignId(_ponyWallCode));
                        ponyType.SetAttributeValue("nominalInsulation", rValue);
                        ponyType.Element("Description")?.SetValue(_ponyWallCode.Label);
                        ponyType.Element("Composite")?.Element("Section")?.SetAttributeValue("nominalRsi", rValue);
                    }
                }
                //Adds the floor header
                if(headerPerim > 0D)
                {
                    bsmt.Add(new XElement("Components"));
                    bsmt.Element("Components").Add(
                        FloorHeader.NewErsJoist(headerHeight.ToString(), headerPerim.ToString(), AssignComponentID().ToString()));

                    if (_floorHeaderCode != null)
                        ApplyCode(bsmt.Element("Components").Element("FloorHeader").Element("Construction").Element("Type"),
                            _floorHeaderCode, includeNominalInsulation: true);
                }
                //Sets the floor above/below frostline attribute based on average depth of basement below grade
                if (depthBelow < 4D)
                    bsmt.Element("Floor").Element("Construction").SetAttributeValue("isBelowFrostline", "false");

                House.Root.Element("House").Element("Components").Add(bsmt);
                XElement? temps = House.Root.Element("House").Element("Temperatures").Element("Basement");
                temps.SetAttributeValue("heated", "true");
            }
            else
            {
                _basementPresent = false;
                SlabOnGrade();
            }
        }

        private void SlabOnGrade()
        {
            bool slabOnGrade = GetCellValue("Calc", "N5")?.ToLower() == "y";
            if (slabOnGrade)
            {
                double area = Math.Round(GetDoubleCellValue("Calc", "F40") * 0.092903, 4);
                double perim = Math.Round(GetDoubleCellValue("Calc", "D40") * 0.3048, 4);

                XElement slab = new XElement("Slab",
                new XAttribute("isExposedSurface", "true"),
                new XAttribute("exposedSurfacePerimeter", perim),
                new XAttribute("id", AssignComponentID()),
                    new XElement("Label", "Slab-on-grade"),
                    new XElement("Configuration",
                        new XAttribute("type", "SCB"),
                        new XAttribute("subtype", "29"),
                        "SCB_29"),
                    new XElement("Floor",
                        new XElement("Construction",
                            new XAttribute("isBelowFrostline", "false"),
                            new XAttribute("hasIntegralFooting", "false"),
                            new XAttribute("heatedFloor", "false"),
                            new XElement("AddedToSlab",
                                new XAttribute("rValue", "0"),
                                new XAttribute("nominalInsulation", "0"),
                                "User specified")),
                        new XElement("Measurements",
                            new XAttribute("isRectangular", "false"),
                            new XAttribute("area", area),
                            new XAttribute("perimeter", perim))),
                    new XElement("Wall",
                        new XElement("RValues",
                            new XAttribute("skirt", "0"),
                            new XAttribute("thermalBreak", "0"))));
                House.Root.Element("House").Element("Components").Add(slab);
            }
        }

        // ── DHW ────────────────────────────────────────────────────────────────

        private void GasDHW()
        {
            foreach (XElement tank in House.Descendants("Components").Descendants("HotWater"))
                tank.Element("Primary").DescendantsAndSelf().Remove();

            bool isPrimary = true;
            if (Convert.ToDouble(GetCellValue("General", "P4")) > 0)
            {
                string dhwMake    = GetCellValue("Summary", "J74");
                string dhwModel   = GetCellValue("Summary", "K74");
                string dhwSize    = GetCellValue("Summary", "K75");
                string dhwEF      = GetCellValue("Summary", "K77");
                bool isUEF        = !GetCellValue("General", "P6").Equals("0");
                string drawPattern = GetCellValue("General", "P6");
                new WaterHeater(dhwMake, dhwModel, dhwEF, dhwSize, false, isPrimary, isUEF, drawPattern, House)
                    .AddTank(_basementPresent);
            }
        }

        private void ElectricDHW()
        {
            bool isPrimary = false;
            if (GetDoubleCellValue("General", "I32") > 0)
            {
                if (GetDoubleCellValue("General", "P4") <= 0)
                    isPrimary = true;

                string electricTankMake   = GetCellValue("Summary", "L74");
                string electricTankModel  = GetCellValue("Summary", "M74");
                string electricTankVolume = GetCellValue("Summary", "M75");
                string electricTankEF     = GetCellValue("Summary", "M77");

                new WaterHeater(electricTankMake, electricTankModel, electricTankEF, electricTankVolume,
                    true, isPrimary, false, "none", House)
                    .AddTank(_basementPresent);
            }
        }
        private void SolarPanels()
        {
            XElement? generation = House.Root.Element("House").Element("Generation");
            if(generation == null)
            {
                generation = new XElement("Generation",
                                new XAttribute("solarReady", "false"),
                                new XAttribute("photovoltaicCapacity", "0"),
                                new XAttribute("batteryStorage", "0"),
                                new XAttribute("id", AssignComponentID()),
                                new XElement("Label", "Generation"));
                House.Root.Add(generation);
            }
            if (double.TryParse(GetCellValue("General", "P33"), out double panelArea) && panelArea > 0)
            {
                int panelCount = 1;
                string model = GetCellValue("General", "P31");
                string make = GetCellValue("General", "P32");
                double efficiency = GetDoubleCellValue("General", "P34");
                double opTemp = GetDoubleCellValue("General", "P35");
                double tempCoeff = GetDoubleCellValue("General", "P36");
                XElement systems = new XElement("PhotovoltaicSystems");

                int startRow = 32;
                int endRow = 36;
                string areaCol = "V";
                string slopeCol = "U";
                string azCol = "T";

                for(int i = startRow; i <= endRow; i++)
                {
                    double area = GetDoubleCellValue("General", areaCol + i);
                    if(area > 0)
                    {
                        SolarPanel panel = new SolarPanel
                        {
                            Area = area,
                            Slope = Math.Round(GetDoubleCellValue("General", slopeCol + i), 4),
                            Azimuth = GetDoubleCellValue("General", azCol + i),
                            Efficiency = efficiency,
                            OperatingTemp = opTemp,
                            TempCoeff = tempCoeff,
                            Manufacturer = make,
                            Model = model,
                            Rank = panelCount++
                        };
                        systems.Add(panel.NewPanel());
                    }
                }
                House.Root.Element("House").Element("Generation").Add(systems);
            }
        }
        private void CheckFireplace()
        {
            bool fpPresent = GetCellValue("Calc", "M57").ToLower() == "y" ? true : false;
            if (fpPresent)
            {
                XElement heatingCooling = House.Root.Element("House").Element("HeatingCooling");
                if(heatingCooling != null)
                {
                    heatingCooling.Add(Fireplace.NewFireplace());
                }
            }
        }
        //Adds a front door with a transom to the first floor wall and a garage door to the garage wall
        private void AddDoors()
        {
            XElement? firstFlr = (from el in House.Descendants("Components").Elements("Wall")
                                  where el.Element("Label").Value.ToString().Contains("1st")
                                  select el).FirstOrDefault();
            XElement? gar = (from el in House.Descendants("Components").Elements("Wall")
                                  where el.Element("Label").Value.ToString().ToLower().Contains("garage")
                                  select el).FirstOrDefault();
            if (firstFlr != null)
            {
                XElement front = Door.FrontDoor(Settings.Default.FrontDoorWidth, Settings.Default.FrontDoorHeight, AssignComponentID());
                if (Settings.Default.FrontTransom)
                {
                    Door.AddTransom(House, front, AssignComponentID(), _nextCodeId++);
                }
                if (firstFlr.Element("Components") == null)
                {
                    firstFlr.Add(new XElement("Components"));
                }
                firstFlr.Element("Components").Add(front);
            }
            if (gar != null)
            {
                XElement garageDoor = Door.PolystyreneDoor(Settings.Default.GarageDoorWidth, Settings.Default.GarageDoorHeight, "Garage", AssignComponentID());
                gar.Element("Components").Add(garageDoor);
            }
        }
        /**
        * Extracts window information from the spreadsheet into a list, then adds the windows and their codes to the house file
        */
        public void ExtractWindows()
        {
            List<Window> windows = new List<Window>();
            bool overhangPres = double.TryParse(GetCellValue("Calc", "M52"), out double overhang);
            int maxWindowRow = Settings.Default.MaxWindowRow;

            for (int i = 2; i <= maxWindowRow; i++)
            {
                string? name = GetCellValue("Windows", "A" + i);
                if (name != null && name != string.Empty && GetCellValue("Windows", "F" + i).ToLower() != "door")
                {
                    bool widthPres = int.TryParse(GetCellValue("Windows", "B" + i), out int width);
                    bool heightPres = int.TryParse(GetCellValue("Windows", "C" + i), out int height);
                    bool uValuePres = double.TryParse(GetCellValue("Windows", "D" + i), out double uValue);
                    bool shgcPres = double.TryParse(GetCellValue("Windows", "E" + i), out double shgc);
                    bool floorPres = int.TryParse(GetCellValue("Windows", "H" + i), out int floor);
                    string operation = GetCellValue("Windows", "G" + i);
                    string direction = GetCellValue("Windows", "I" + i);

                    if (!widthPres || !heightPres || !uValuePres || !shgcPres || !floorPres)
                        continue;

                    name = $"{name}-{operation}";
                    Window window = new(name, width, height, uValue, shgc, floor, direction, overhangPres ? overhang : 0, AssignComponentID());
                    windows.Add(window);
                    window.codeId = CodeTools.FindWindowCode(House, window);
                    window.AddWindow(House);
                }
            }
        }
        /**
        * Removes all windows that aren't a part of door assemblies.
        */
        public void RemoveWindows()
        {
            List<XElement>? windows = House?.Root?.Element("House")?.Descendants("Window").
                            Where(el => !el.Ancestors("Door").Any() && el != null).
                            ToList();
            foreach (XElement window in windows)
            {
                window.Remove();
            }
        }

        // ── Code assignment ────────────────────────────────────────────────────
        // Codes are applied to each component as it is created (see ProcessWalls,
        // AddCeiling, NewFloor, CheckBasement). BuildCodesSection() then collects the
        // used codes into the house file's <Codes> block.

        private void ApplyCode(XElement typeEl, CodeEntry code, bool includeNominalInsulation)
        {
            string assignedId = GetOrAssignId(code);
            string rValue     = code.NominalRValue.ToString();

            typeEl.SetAttributeValue("idref", assignedId);
            typeEl.SetAttributeValue("rValue", rValue);
            if (includeNominalInsulation)
                typeEl.SetAttributeValue("nominalInsulation", rValue);
            typeEl.Value = code.Label;
        }

        private string GetOrAssignId(CodeEntry code)
        {
            if (!_assignedIds.TryGetValue(code.Label, out string? id))
            {
                id = $"Code {_nextCodeId++}";
                _assignedIds[code.Label] = id;
                _usedCodes.Add(code);
            }
            return id;
        }

        private void BuildCodesSection()
        {
            XElement? codesEl = House.Root?.Element("Codes");
            if (codesEl == null)
            {
                codesEl = new XElement("Codes");
                House.Root?.Add(codesEl);
            }

            var sectionsToRebuild = new HashSet<string> { "Wall", "Ceiling", "CeilingFlat", "Floor", "FloorHeader", "FloorsAbove", "BasementWall", "CrawlspaceWall" };
            foreach (string section in sectionsToRebuild)
                codesEl.Element(section)?.Remove();

            foreach (var sectionGroup in _usedCodes.GroupBy(c => c.CodSection))
            {
                XElement sectionEl = new XElement(sectionGroup.Key);

                // Emit each code under the wrapper it came from (UserDefined vs Favorite):
                // their schema types differ, so mixing them up produces an invalid file.
                foreach (var wrapperGroup in sectionGroup.GroupBy(c => c.Wrapper))
                {
                    sectionEl.Add(new XElement(wrapperGroup.Key,
                        wrapperGroup.Select(code =>
                        {
                            XElement codeEl = new XElement(code.Element);
                            codeEl.SetAttributeValue("id", _assignedIds[code.Label]);
                            return codeEl;
                        })));
                }

                codesEl.Add(sectionEl);
            }
        }
    }
}
