using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotPort.Models
{
    public class Energuide
    {
        public XDocument House { get; }

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

        // Tracks codes used during generation: label → assigned "Code N" id
        private readonly Dictionary<string, string> _assignedIds = new();
        private readonly List<CodeEntry> _usedCodes = new();
        private int _nextId = 1;

        public Energuide(
            XDocument template,
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

        public void AssignCodes()
        {
            AssignWallCodes();
            AssignFloorHeaderCodes();
            AssignCeilingCodes();
            AssignFloorCodes();
            BuildCodesSection();
        }

        // ── Walls ──────────────────────────────────────────────────────────────

        private void AssignWallCodes()
        {
            foreach (XElement wall in House.Descendants("Components").Elements("Wall"))
            {
                string label = wall.Element("Label")?.Value ?? string.Empty;
                XElement? typeEl = wall.Element("Construction")?.Element("Type");
                if (typeEl == null) continue;

                string currentLabel = typeEl.Value;
                CodeEntry? code = SelectWallCode(label, currentLabel);
                if (code != null)
                    ApplyCode(typeEl, code, includeNominalInsulation: true);
            }
        }

        private CodeEntry? SelectWallCode(string wallLabel, string currentTypeLabel)
        {
            if (wallLabel.Contains("Tall", StringComparison.OrdinalIgnoreCase))
                return _tallWallCode;
            if (currentTypeLabel.Contains("NoClad", StringComparison.OrdinalIgnoreCase))
                return _noCladWallCode;
            return _mainWallCode;
        }

        // ── Floor headers ──────────────────────────────────────────────────────

        private void AssignFloorHeaderCodes()
        {
            foreach (XElement fh in House.Descendants("FloorHeader"))
            {
                XElement? typeEl = fh.Element("Construction")?.Element("Type");
                if (typeEl != null && _floorHeaderCode != null)
                    ApplyCode(typeEl, _floorHeaderCode, includeNominalInsulation: true);
            }
        }

        // ── Ceilings ───────────────────────────────────────────────────────────

        private void AssignCeilingCodes()
        {
            foreach (XElement ceiling in House.Descendants("Ceiling"))
            {
                string typeCode = ceiling.Element("Construction")
                    ?.Element("Type")
                    ?.Attribute("code")?.Value ?? string.Empty;

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

        // ── Floors ─────────────────────────────────────────────────────────────

        private void AssignFloorCodes()
        {
            // Only above-grade <Floor> elements directly under <Components>
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

        // ── Code assignment helper ─────────────────────────────────────────────

        private void ApplyCode(XElement typeEl, CodeEntry code, bool includeNominalInsulation)
        {
            string assignedId = GetOrAssignId(code);
            string rValue = code.NominalRValue.ToString();

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

        // ── <Codes> section builder ────────────────────────────────────────────

        private void BuildCodesSection()
        {
            XElement? codesEl = House.Root?.Element("Codes");
            if (codesEl == null)
            {
                codesEl = new XElement("Codes");
                House.Root?.Add(codesEl);
            }

            // Remove only the sections we're replacing; leave Window etc. intact
            var sectionsToRebuild = new HashSet<string> { "Wall", "Ceiling", "CeilingFlat", "Floor", "FloorHeader" };
            foreach (string section in sectionsToRebuild)
                codesEl.Element(section)?.Remove();

            // Group used codes by their COD section and write each group
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
