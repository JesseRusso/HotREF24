using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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

        private readonly Dictionary<string, string> _assignedIds = new();
        private readonly List<CodeEntry> _usedCodes = new();
        private int _nextId = 1;

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
            CodeEntry? garageFloorCode)
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
        }

        // ── Entry points ───────────────────────────────────────────────────────

        public void Generate()
        {
            BuildHouse();
            AssignCodes();
        }

        public void BuildHouse()
        {
            GetBuilder();
            FindID();
            InitNextCodeId();
            CityCheck();
            ChangeCityWeather();
            ChangeSpecs();
            ChangeEquipment();
            CheckAC();
            ProcessWalls();
            CheckCeilings();
            ExtraCeilings();
            CheckVaults();
            ChangeFloors();
            ExtraFloors();
            ChangeBasment();
            GasDHW();
            ElectricDHW();
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
        }

        // ── Infrastructure ─────────────────────────────────────────────────────

        private void GetBuilder()
        {
            Builder = House.Descendants("File").Descendants("BuilderName")
                          .Where(el => !el.IsEmpty)
                          .First().Value;
        }

        private void FindID()
        {
            var ids = House.Descendants("House").Descendants().Attributes("id")
                          .Where(a => a.Value != null)
                          .Select(a => int.Parse(a.Value))
                          .ToList();
            MaxID = ids.Max() + 1;
        }

        private void InitNextCodeId()
        {
            // Sections we will rebuild — codes in these will be replaced, so their ids don't matter.
            var rebuilt = new HashSet<string> { "Wall", "Ceiling", "CeilingFlat", "Floor", "FloorHeader" };

            // Scan sections we keep (Window, Door, …) for the highest "Code N" id already in use.
            int max = 0;
            XElement? codesEl = House.Root?.Element("Codes");
            if (codesEl != null)
            {
                foreach (XElement section in codesEl.Elements().Where(e => !rebuilt.Contains(e.Name.LocalName)))
                {
                    foreach (XAttribute attr in section.Descendants().Attributes("id"))
                    {
                        if (attr.Value.StartsWith("Code ") &&
                            int.TryParse(attr.Value.AsSpan(5), out int n))
                            max = Math.Max(max, n);
                    }
                }
            }
            _nextId = max + 1;
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

            SetDate();
        }

        private void SetDate()
        {
            XElement evalDate = House.Element("HouseFile").Element("ProgramInformation").Element("File");
            evalDate.SetAttributeValue("evaluationDate", DateTime.UtcNow.Date.ToString("yyyy-MM-dd"));
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

        private void ChangeEquipment()
        {
            ChangeHRV();
            ChangeFurnace();
        }

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
                }
            }
        }

        private void ChangeFurnace()
        {
            string furnaceModel = GetCellValue("Summary", "B78");
            double furnaceBtus  = GetDoubleCellValue("General", "C4");

            if (GetDoubleCellValue("General", "B4") > 0)
                furnaceModel += $" & {GetCellValue("Summary", "B79")}";

            foreach (XElement furn in House.Descendants("Furnace"))
            {
                furn.Element("Specifications")?.SetAttributeValue("efficiency", GetCellValue("General", "A6"));
                furn.Element("Specifications")?.Element("OutputCapacity")
                    ?.SetAttributeValue("value", Math.Round(furnaceBtus * 0.00029307107, 5).ToString());
                furn.Element("EquipmentInformation").Element("Manufacturer").SetValue(GetCellValue("Summary", "A78"));
                furn.Element("EquipmentInformation").Element("Model").SetValue(furnaceModel);
            }
            foreach (XElement fan in House.Descendants("HeatingCooling"))
            {
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("low", GetCellValue("General", "E4"));
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("high", GetCellValue("General", "D4"));
            }
        }

        private void CheckAC()
        {
            string ACBtus = GetCellValue("General", "Q26");
            if (!int.TryParse(ACBtus, out int cap) || cap <= 0) return;

            House?.Element("HouseFile")?.Element("House")?.Element("Temperatures")?.Element("Basement")
                 ?.SetAttributeValue("cooled", "true");

            XElement? type2 = House?.Descendants("HeatingCooling").Descendants("Type2").FirstOrDefault();
            type2?.Add(
                new XElement("AirConditioning",
                    new XElement("EquipmentInformation",
                        new XAttribute("energystar", "false"),
                        new XElement("Manufacturer", GetCellValue("General", "N25")),
                        new XElement("Model", GetCellValue("General", "P25"))),
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
                            new XAttribute("value", GetCellValue("General", "O26")))),
                    new XElement("CoolingParameters",
                        new XAttribute("sensibleHeatRatio", "0.76"),
                        new XAttribute("openableWindowArea", "0"),
                        new XElement("FansAndPump",
                            new XAttribute("flowRate", "123"),
                            new XAttribute("hasEnergyEfficientMotor", "false"),
                            new XElement("Mode", new XAttribute("code", "1")),
                            new XElement("Power", new XAttribute("isCalculated", "true"))))));
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
                double fhPerim       = GetDoubleCellValue("Calc", $"J{row}");
                double fhHeight      = GetDoubleCellValue("Calc", $"K{row}");
                string spacing       = GetCellValue("Calc", $"B{row}") ?? string.Empty;
                string siding        = GetCellValue("Calc", $"D{row}") ?? string.Empty;

                AddNewWall(name, corners, intersections, height, perimeter);

                XElement? newWall = House.Descendants("Wall")
                    .FirstOrDefault(w => w.Attribute("id")?.Value == (MaxID - 1).ToString());
                XElement? wallTypeEl = newWall?.Element("Construction")?.Element("Type");
                CodeEntry? wallCode = WallCodeFromSpecs(spacing, siding);
                if (wallTypeEl != null && wallCode != null)
                    ApplyCode(wallTypeEl, wallCode, includeNominalInsulation: true);

                if (fhPerim > 0)
                {
                    if (newWall?.Element("Components") == null)
                        newWall!.Add(new XElement("Components"));
                    newWall?.Element("Components")?.Add(
                        FloorHeader.NewJoist(fhHeight.ToString(), "0", fhPerim.ToString(), MaxID.ToString()));
                    MaxID++;
                }
            }

            SetHouseType();
            SetStoreys();
        }

        private CodeEntry? WallCodeFromSpecs(string spacing, string siding) =>
            siding.Equals("n/a", StringComparison.OrdinalIgnoreCase) ? _noCladWallCode :
            spacing.Contains("12")                                    ? _tallWallCode :
            _mainWallCode;

        private void AddNewWall(string name, string corners, string intersections, double height, double perim)
        {
            House.Descendants("Components").First().Add(
                new XElement("Wall",
                    new XAttribute("adjacentEnclosedSpace", "false"),
                    new XAttribute("id", MaxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XAttribute("corners", corners),
                        new XAttribute("intersections", intersections),
                        new XElement("Type", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("height",    Math.Round(height * 0.3048, 4).ToString()),
                        new XAttribute("perimeter", Math.Round(perim  * 0.3048, 4).ToString())),
                    new XElement("FacingDirection",
                        new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O"))));
            MaxID++;
        }

        // ── Ceilings ───────────────────────────────────────────────────────────

        private void CheckCeilings()
        {
            // Remove all template ceiling placeholders; actual ceilings are added by ExtraCeilings/CheckVaults
            House.Descendants("Ceiling")
                 .Where(c => c.Element("Label")?.Value.Contains("2nd") == true)
                 .ToList().ForEach(c => c.Remove());

            House.Descendants("Components").Descendants("Ceiling")
                 .Where(c => new[] { "4", "5", "6" }.Contains(
                     c.Element("Construction")?.Element("Type")?.Attribute("code")?.Value))
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

            for (int row = 10; row <= 17; row++)
            {
                string cellValue = GetCellValue("Calc", $"M{row}");
                if (string.IsNullOrEmpty(cellValue) || !double.TryParse(cellValue, out double area) || area <= 0)
                    continue;

                double length = GetDoubleCellValue("Calc", $"L{row}");
                string name   = GetCellValue("Calc", $"I{row}");
                string type   = GetCellValue("Calc", $"AD{row}");
                string slope  = GetCellValue("Calc", $"N{row}");

                AddCeiling(name, type, area, length, slope, heel, isVault: true);
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

            XElement comp = House.Descendants("Components").First();
            comp.Add(
                new XElement("Ceiling",
                    new XAttribute("id", MaxID),
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
                            new XElement("French",  slopeFr)))));
            MaxID++;
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

        private void ChangeFloors()
        {
            double garFlrArea = GetDoubleCellValue("Calc", "P21");
            bool garPresent   = garFlrArea > 0;

            XElement? garFlr = House.Descendants("Floor")
                .FirstOrDefault(el => el.Element("Label").Value.ToLower().Contains("garage"));

            XElement? floor = House.Descendants("Floor")
                .FirstOrDefault(el => el.Element("Label").Value.ToLower().Contains("cant"));

            floor?.Remove();

            if (garFlr != null)
            {
                if (garPresent)
                {
                    double garFlrLength = GetDoubleCellValue("Calc", "O21");
                    garFlr.Element("Measurements")?.SetAttributeValue("area",   Math.Round(garFlrArea   * 0.092903, 4));
                    garFlr.Element("Measurements")?.SetAttributeValue("length", Math.Round(garFlrLength * 0.3048,   4));
                    garFlr.Element("Label")?.SetValue(GetCellValue("Calc", "L21"));
                }
                else
                {
                    garFlr.Remove();
                }
            }
        }

        private void ExtraFloors()
        {
            for (int row = 22; row <= 34; row++)
            {
                string cellValue = GetCellValue("Calc", $"P{row}");
                if (string.IsNullOrEmpty(cellValue) || !double.TryParse(cellValue, out double area) || area <= 0)
                    continue;

                double length = GetDoubleCellValue("Calc", $"O{row}");
                string name   = GetCellValue("Calc", $"L{row}");
                NewFloor(name, length, area);
            }
        }

        private void NewFloor(string name, double length, double area)
        {
            XElement comp = House.Descendants("Components").First();
            comp.Add(
                new XElement("Floor",
                    new XAttribute("adjacentEnclosedSpace", "false"),
                    new XAttribute("id", MaxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XElement("Type", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("area",   Math.Round(area   * 0.092903, 4).ToString()),
                        new XAttribute("length", Math.Round(length * 0.3048,   4).ToString()))));
            MaxID++;
        }

        // ── Basement ───────────────────────────────────────────────────────────

        private void ChangeBasment()
        {
            bool basement       = GetCellValue("Calc", "N2")?.ToLower() == "y";
            bool bsmtUnder4Feet = GetCellValue("Calc", "N4")?.ToLower() == "y";
            bool slabOnGrade    = GetCellValue("Calc", "N5")?.ToLower() == "y";

            if (!basement && !bsmtUnder4Feet)
                _basementPresent = false;

            if (basement)
            {
                XElement over4 = House.Descendants("Components").Descendants("Basement")
                    .First(el => el.Element("Label").Value.Contains(">"));

                over4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(Convert.ToDouble(GetCellValue("Calc", "E38")) * 0.3048, 4));
                over4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(GetDoubleCellValue("Calc", "D38") * 0.3048, 4));
                over4.Element("Floor").Element("Measurements").SetAttributeValue("area",      Math.Round(GetDoubleCellValue("Calc", "F38") * 0.092903, 4));
                over4.Element("Wall").Element("Measurements").SetAttributeValue("height",     Math.Round(GetDoubleCellValue("Calc", "G38") * 0.3048, 4));
                over4.Element("Wall").Element("Measurements").SetAttributeValue("depth",      Math.Round(GetDoubleCellValue("Calc", "H38") * 0.3048, 4));
                over4.Element("Wall").Element("Construction").SetAttributeValue("corners",    GetCellValue("Calc", "J38"));
                over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height",    Math.Round(GetDoubleCellValue("Calc", "K38") * 0.3048, 4));
                over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(GetDoubleCellValue("Calc", "L38") * 0.3048, 4));

                string pony = GetCellValue("Calc", "I38");
                if (double.TryParse(pony, out double ponyHeight) && ponyHeight > 0D)
                {
                    over4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                    over4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(ponyHeight * 0.3048, 4));
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

            House.Descendants("Components").Descendants("Basement")
                 .FirstOrDefault(x => x.Element("Label").Value.Contains(">"))
                 ?.Remove();

            Under4Bsmt(bsmtUnder4Feet);
            SlabOnGrade(slabOnGrade);
        }

        private void Under4Bsmt(bool under4Present)
        {
            XElement? under4 = House.Descendants("Components").Descendants("Basement")
                .FirstOrDefault(el => el.Element("Label").Value.Contains("<"));

            if (under4 == null) return;

            if (under4Present)
            {
                under4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(GetDoubleCellValue("Calc", "E39") * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(GetDoubleCellValue("Calc", "D39") * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("area",      Math.Round(GetDoubleCellValue("Calc", "F39") * 0.092903, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("height",     Math.Round(GetDoubleCellValue("Calc", "G39") * 0.3048, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("depth",      Math.Round(GetDoubleCellValue("Calc", "H39") * 0.3048, 4));
                under4.Element("Wall").Element("Construction").SetAttributeValue("corners",    GetCellValue("Calc", "J39"));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height",    Math.Round(GetDoubleCellValue("Calc", "K39") * 0.3048, 4));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(GetDoubleCellValue("Calc", "L39") * 0.3048, 4));

                string pony = GetCellValue("Calc", "I39");
                if (double.TryParse(pony, out double ponyHeight) && ponyHeight > 0D)
                {
                    under4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                    under4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(ponyHeight * 0.3048, 4));
                    under4.Element("Wall").Element("Construction").Add(
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
            }
            else
            {
                under4.Remove();
            }
        }

        private void SlabOnGrade(bool isSlabPresent)
        {
            XElement? slab = House.Descendants("Components").Descendants("Slab")
                .FirstOrDefault(el => el.Element("Label").Value.Contains("Slab"));

            if (slab == null) return;

            if (isSlabPresent)
            {
                slab.SetAttributeValue("exposedSurfacePerimeter", Math.Round(GetDoubleCellValue("Calc", "E40") * 0.3048, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("area",      Math.Round(GetDoubleCellValue("Calc", "F40") * 0.092903, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(GetDoubleCellValue("Calc", "D40") * 0.3048, 4));
            }
            else
            {
                slab.Remove();
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

        // ── Code assignment ────────────────────────────────────────────────────

        public void AssignCodes()
        {
            AssignFloorHeaderCodes();
            AssignCeilingCodes();
            AssignFloorCodes();
            BuildCodesSection();
        }

        private void AssignFloorHeaderCodes()
        {
            foreach (XElement fh in House.Descendants("FloorHeader"))
            {
                XElement? typeEl = fh.Element("Construction")?.Element("Type");
                if (typeEl != null && _floorHeaderCode != null)
                    ApplyCode(typeEl, _floorHeaderCode, includeNominalInsulation: true);
            }
        }

        private void AssignCeilingCodes()
        {
            foreach (XElement ceiling in House.Descendants("Ceiling"))
            {
                string typeCode = ceiling.Element("Construction")
                    ?.Element("Type")?.Attribute("code")?.Value ?? string.Empty;

                XElement? ceilingTypeEl = ceiling.Element("Construction")?.Element("CeilingType");
                if (ceilingTypeEl == null) continue;

                CodeEntry? code = typeCode switch
                {
                    "2" or "3" => _ceilingCode,
                    "6"        => _vaultCode,
                    "4"        => _cathedralCode,
                    "5"        => _flatCode,
                    _          => null
                };

                if (code != null)
                    ApplyCode(ceilingTypeEl, code, includeNominalInsulation: true);
            }
        }

        private void AssignFloorCodes()
        {
            foreach (XElement floor in House.Descendants("Components").Elements("Floor"))
            {
                string label = floor.Element("Label")?.Value ?? string.Empty;
                XElement? typeEl = floor.Element("Construction")?.Element("Type");
                if (typeEl == null) continue;

                CodeEntry? code = label.Contains("Gar", StringComparison.OrdinalIgnoreCase)
                    ? _garageFloorCode
                    : _exposedFloorCode;

                if (code != null)
                    ApplyCode(typeEl, code, includeNominalInsulation: true);
            }
        }

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
                id = $"Code {_nextId++}";
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

            var sectionsToRebuild = new HashSet<string> { "Wall", "Ceiling", "CeilingFlat", "Floor", "FloorHeader" };
            foreach (string section in sectionsToRebuild)
                codesEl.Element(section)?.Remove();

            foreach (var group in _usedCodes.GroupBy(c => c.CodSection))
            {
                XElement sectionEl = new XElement(group.Key,
                    new XElement("UserDefined",
                        group.Select(code =>
                        {
                            XElement codeEl = new XElement(code.Element);
                            codeEl.SetAttributeValue("id", _assignedIds[code.Label]);
                            return codeEl;
                        })));
                codesEl.Add(sectionEl);
            }
        }
    }
}
