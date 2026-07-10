using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotPort.Models
{
    public class CodeEntry
    {
        public string Id { get; }
        public string Label { get; }
        public double NominalRValue { get; }
        public XElement Element { get; }
        public string CodSection { get; }

        public CodeEntry(XElement code, string codSection)
        {
            Id = code.Attribute("id")?.Value ?? string.Empty;
            Label = code.Element("Label")?.Value ?? string.Empty;
            NominalRValue = double.TryParse(code.Attribute("nominalRValue")?.Value, out double r) ? r : 0;
            Element = code;
            CodSection = codSection;
        }

        public override string ToString() => Label;
    }

    public class CodeLibrary
    {
        private readonly XDocument _cod;

        public string FilePath { get; }

        public CodeLibrary(string filePath)
        {
            FilePath = filePath;
            _cod = XDocument.Load(filePath);
        }

        public IReadOnlyList<CodeEntry> GetWallCodes() =>
            CodesUnder("Wall");

        public IReadOnlyList<CodeEntry> GetFloorHeaderCodes() =>
            CodesUnder("FloorHeader");

        public IReadOnlyList<CodeEntry> GetCeilingCodes() =>
            CodesUnder("Ceiling")
                .Where(e => !e.Label.Contains("Vault") && !e.Label.Contains("Vlt"))
                .ToList();

        public IReadOnlyList<CodeEntry> GetVaultCodes() =>
            CodesUnder("Ceiling")
                .Where(e => e.Label.Contains("Vault") || e.Label.Contains("Vlt"))
                .ToList();

        public IReadOnlyList<CodeEntry> GetCathedralCodes() =>
            CodesUnder("CeilingFlat")
                .Where(e => e.Label.Contains("Cath"))
                .ToList();

        public IReadOnlyList<CodeEntry> GetFlatCodes() =>
            CodesUnder("CeilingFlat")
                .Where(e => e.Label.Contains("Flat"))
                .ToList();

        public IReadOnlyList<CodeEntry> GetFloorCodes() =>
            CodesUnder("Floor");

        public IReadOnlyList<CodeEntry> ExposedFloorCodes() =>
            GetFloorCodes().Where(e => !e.Label.Contains("Gar")).ToList();

        public IReadOnlyList<CodeEntry> GarageFloorCodes() =>
            GetFloorCodes().Where(e => e.Label.Contains("Gar")).ToList();

        // Infer wall code from stud spacing and siding type strings
        public CodeEntry? InferWallCode(string spacing, string siding) =>
            BestMatch(GetWallCodes(), spacing, siding);

        // Floor headers share siding with walls but have no spacing
        public CodeEntry? InferFloorHeaderCode(string siding) =>
            BestMatch(GetFloorHeaderCodes(), siding);

        public CodeEntry? InferCeilingCode() =>
            GetCeilingCodes().FirstOrDefault();

        public CodeEntry? InferVaultCode() =>
            GetVaultCodes().FirstOrDefault();

        public CodeEntry? InferCathedralCode() =>
            GetCathedralCodes().FirstOrDefault();

        public CodeEntry? InferFlatCode() =>
            GetFlatCodes().FirstOrDefault();

        public CodeEntry? InferExposedFloorCode() =>
            GetFloorCodes().FirstOrDefault(e => !e.Label.Contains("Gar"));

        public CodeEntry? InferGarageFloorCode() =>
            GetFloorCodes().FirstOrDefault(e => e.Label.Contains("Gar"));

        // Returns all UserDefined codes under the named component section
        private IReadOnlyList<CodeEntry> CodesUnder(string sectionName) =>
            _cod.Root?
                .Element(sectionName)?
                .Element("UserDefined")?
                .Elements("Code")
                .Select(e => new CodeEntry(e, sectionName))
                .ToList()
            ?? new List<CodeEntry>();

        // Picks the shortest label that contains all provided terms (case-insensitive, whitespace-normalized)
        private static CodeEntry? BestMatch(IReadOnlyList<CodeEntry> codes, params string[] terms)
        {
            // Normalize terms and label by removing spaces so "16 OC" matches "16OC"
            string[] normalizedTerms = terms
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Replace(" ", string.Empty))
                .ToArray();

            return codes
                .Where(e =>
                {
                    string normalizedLabel = e.Label.Replace(" ", string.Empty);
                    return normalizedTerms.All(t =>
                        normalizedLabel.Contains(t, System.StringComparison.OrdinalIgnoreCase));
                })
                .OrderBy(e => e.Label.Length)
                .FirstOrDefault();
        }
    }
}
