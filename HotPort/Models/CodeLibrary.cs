using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace HotPort.Models
{
    public record CodeRef(string Idref, string RValue, string Label);

    public enum WallSiding { Unknown, NoClad, Vinyl, Hardie, Stucco }

    public class CodeEntry
    {
        public string Id { get; }
        public string Label { get; }
        public double NominalRValue { get; }
        public XElement Element { get; }
        public string CodSection { get; }
        // The COD block this code lives under: "UserDefined" or "Favorite".
        // Determines which wrapper it must be written under in the house file's <Codes>.
        public string Wrapper { get; }
        // Nominal stud spacing (inches OC) parsed from the framing layer; null if not derivable.
        public int? StudSpacingOC { get; }
        // Exterior finish parsed from the siding layer; NoClad when absent.
        public WallSiding Siding { get; }

        public CodeEntry(XElement code, string codSection)
        {
            Id = code.Attribute("id")?.Value ?? string.Empty;
            Label = code.Element("Label")?.Value ?? string.Empty;
            NominalRValue = double.TryParse(code.Attribute("nominalRValue")?.Value, out double r) ? r : 0;
            Element = code;
            CodSection = codSection;
            Wrapper = code.Parent?.Name.LocalName ?? "UserDefined";
            StudSpacingOC = ParseStudSpacingOC(code);
            Siding = ParseSiding(code);
        }

        // Parses nominal stud spacing (inches OC), handling both code layer formats.
        private static int? ParseStudSpacingOC(XElement code)
        {
            // UserDefined format: WoodFraming/Framing/@spacing is metric (mm).
            XElement? framing = code.Descendants("Framing")
                .FirstOrDefault(f => f.Attribute("spacing") != null);
            if (framing != null &&
                double.TryParse(framing.Attribute("spacing")!.Value, out double mm) && mm > 0)
                return (int)Math.Round(mm / 25.4);

            // Standard/Favorite format: <Spacing> layer's code.
            return code.Descendants("Spacing").FirstOrDefault()?.Attribute("code")?.Value switch
            {
                "0" => 12,
                "1" => 16,
                "2" => 19,
                "3" => 24,
                _   => (int?)null
            };
        }

        // Parses the exterior finish, handling both code layer formats.
        private static WallSiding ParseSiding(XElement code)
        {
            // UserDefined format: siding is a <ContinuousMedium> whose Material/Category
            // code is 4 (exterior siding) or 2 (masonry/stucco).
            var continuousMedia = code.Descendants("ContinuousMedium").ToList();
            if (continuousMedia.Count > 0)
            {
                XElement? cm = continuousMedia.LastOrDefault(m =>
                {
                    string? cat = m.Element("Material")?.Element("Category")?.Attribute("code")?.Value;
                    return cat == "4" || cat == "2";
                });
                if (cm == null) return WallSiding.NoClad;

                string? category = cm.Element("Material")?.Element("Category")?.Attribute("code")?.Value;
                string? type = cm.Element("Material")?.Element("Type")?.Attribute("code")?.Value;

                if (category == "2") return WallSiding.Stucco;
                // category == "4" (exterior siding): Type distinguishes the product
                return type switch
                {
                    "4" => WallSiding.Vinyl,
                    "2" => WallSiding.Hardie,
                    _   => WallSiding.Unknown
                };
            }

            // Standard/Favorite format: siding is the <Exterior> layer's code.
            return code.Descendants("Exterior").FirstOrDefault()?.Attribute("code")?.Value switch
            {
                "0" => WallSiding.NoClad,
                "1" => WallSiding.Hardie,  // Wood (lapped)
                "2" => WallSiding.Vinyl,   // Hollow metal/vinyl cladding
                "6" => WallSiding.Stucco,
                _   => WallSiding.Unknown
            };
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

        public IReadOnlyList<CodeEntry> GetCeilingCodes()
        {
            var codes = CodesUnder("Ceiling")
                .Where(e => !e.Label.Contains("Vault") && !e.Label.Contains("Vlt"))
                .ToList();
            return codes;
        }


        public IReadOnlyList<CodeEntry> GetVaultCodes()
        {
           var vaults = CodesUnder("Ceiling")
                .Where(e => e.Label.Contains("Vault") || e.Label.Contains("Vlt"))
                .ToList();
            return vaults.Count > 0 ? vaults : GetCeilingCodes();
        }


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

        public IReadOnlyList<CodeEntry> GetFloorsAboveCodes() =>
            CodesUnder("FloorsAbove");

        // Interior basement wall insulation codes
        public IReadOnlyList<CodeEntry> GetInteriorWallCodes() =>
            CodesUnder("BasementWall");

        // Pony wall codes live under the CrawlspaceWall section
        public IReadOnlyList<CodeEntry> GetPonyWallCodes() =>
            CodesUnder("CrawlspaceWall");

        public IReadOnlyList<CodeEntry> ExposedFloorCodes() =>
            GetFloorCodes().Where(e => !e.Label.Contains("Gar")).ToList();

        public IReadOnlyList<CodeEntry> GarageFloorCodes() =>
            GetFloorCodes().Where(e => e.Label.Contains("Gar")).ToList();

        // Infer wall code from stud spacing and siding type strings (spreadsheet values).
        // Prefer a structured match on the code's parsed framing spacing + siding layer;
        // fall back to label matching when the request or the codes aren't structured.
        public CodeEntry? InferWallCode(string spacing, string siding)
        {
            var walls = GetWallCodes();

            int? wantOC = ParseRequestedSpacingOC(spacing);
            WallSiding wantSiding = ParseRequestedSiding(siding);

            if (wantOC != null && wantSiding != WallSiding.Unknown)
            {
                CodeEntry? structured = walls
                    .Where(c => c.StudSpacingOC == wantOC && c.Siding == wantSiding)
                    .OrderBy(c => c.Label.Length)  // prefer the base variant over ZLL/DW/etc.
                    .FirstOrDefault();
                if (structured != null)
                    return structured;
            }

            return BestMatch(walls, spacing, siding);
        }

        // "24 OC" -> 24, "16 OC" -> 16; null if no number present.
        private static int? ParseRequestedSpacingOC(string spacing)
        {
            if (string.IsNullOrEmpty(spacing)) return null;
            Match m = Regex.Match(spacing, @"\d+");
            return m.Success && int.TryParse(m.Value, out int oc) ? oc : (int?)null;
        }

        private static WallSiding ParseRequestedSiding(string siding)
        {
            if (string.IsNullOrWhiteSpace(siding)) return WallSiding.Unknown;
            return siding.Replace(" ", string.Empty).ToLowerInvariant() switch
            {
                "vinyl"  => WallSiding.Vinyl,
                "hardie" => WallSiding.Hardie,
                "stucco" => WallSiding.Stucco,
                "noclad" or "n/a" or "na" => WallSiding.NoClad,
                _ => WallSiding.Unknown
            };
        }

        // Floor headers share siding with walls but have no spacing
        public CodeEntry? InferFloorHeaderCode(string siding)
        {
            var codes = GetFloorHeaderCodes();
            return BestMatch(codes, siding) ?? codes.FirstOrDefault();
        }

        public CodeEntry? InferCeilingCode() =>
            GetCeilingCodes().OrderByDescending(e => e.NominalRValue).FirstOrDefault();

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

        public CodeEntry? InferFloorsAboveCode() =>
            GetFloorsAboveCodes().FirstOrDefault();

        public CodeEntry? InferInteriorWallCode() =>
            GetInteriorWallCodes().OrderByDescending(e => e.NominalRValue).LastOrDefault();

        // Pony walls carry the above-grade siding, so match the main wall's siding when possible.
        public CodeEntry? InferPonyWallCode(string siding)
        {
            var codes = GetPonyWallCodes();
            WallSiding wantSiding = ParseRequestedSiding(siding);

            if (wantSiding != WallSiding.Unknown)
            {
                CodeEntry? structured = codes
                    .Where(c => c.Siding == wantSiding)
                    .OrderBy(c => c.Label.Length)
                    .FirstOrDefault();
                if (structured != null)
                    return structured;
            }

            return BestMatch(codes, siding) ?? codes.FirstOrDefault();
        }

        // Returns all codes under the named component section, from both the
        // <UserDefined> and <Favorite> blocks (a section may use either or both)
        private IReadOnlyList<CodeEntry> CodesUnder(string sectionName) =>
            _cod.Root?
                .Element(sectionName)?
                .Elements()
                .Where(b => b.Name.LocalName is "UserDefined" or "Favorite")
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
