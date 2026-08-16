using System.Linq;
using System.Xml.Linq;

namespace HotPort.Models
{
    public class Foundation
    {
        public int Id { get; init; }
        public string Label { get; init; } = "Basement";
        public double ExposedSurfacePerimeter { get; init; }

        // Floor
        public double FloorPerimeter { get; init; }
        public double FloorArea { get; init; }

        // Wall
        public string WallCorners { get; init; } = "0";
        public double WallHeight { get; init; }
        public double WallDepth { get; init; }
        public double PonyWallHeight { get; init; }

        // Builds the <Basement> element for the house's <Components> block.
        // The caller appends it and writes the pony wall / rim joist floor header.
        public XElement AddBasement()
        {
            return new XElement("Basement",
                new XAttribute("isExposedSurface", "true"),
                new XAttribute("exposedSurfacePerimeter", ExposedSurfacePerimeter),
                new XAttribute("id", Id),
                new XElement("Label", Label),
                new XElement("Configuration", "BCCB_4",
                    new XAttribute("type", "BCCB"),
                    new XAttribute("subtype", "4"),
                    new XAttribute("overlap", "0")),
                new XElement("OpeningUpstairs",
                    new XAttribute("code", "1"),
                    new XAttribute("value", "1.56"),
                    new XElement("English", "Standard door - open"),
                    new XElement("French", "Porte standard - ouverte")),
                new XElement("RoomType",
                    new XAttribute("code", "6"),
                    new XElement("English", "Utility Room"),
                    new XElement("French", "Pièce Utilitaire")),
                new XElement("Floor",
                    new XElement("Construction",
                        new XAttribute("isBelowFrostline", "true"),
                        new XAttribute("hasIntegralFooting", "false"),
                        new XAttribute("heatedFloor", "false"),
                        new XElement("AddedToSlab", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0")),
                        new XElement("FloorsAbove", "User specified",
                            new XAttribute("rValue", "0"),
                            new XAttribute("nominalInsulation", "0"))),
                    new XElement("Measurements",
                        new XAttribute("isRectangular", "false"),
                        new XAttribute("perimeter", FloorPerimeter),
                        new XAttribute("area", FloorArea))),
                new XElement("Wall",
                    new XAttribute("hasPonyWall", PonyWallHeight > 0 ? "true" : "false"),
                    new XElement("Construction",
                        new XAttribute("corners", WallCorners),
                        new XElement("InteriorAddedInsulation",
                            new XAttribute("nominalInsulation", "0"),
                            new XElement("Description", "User specified"),
                            new XElement("Composite",
                                new XElement("Section",
                                    new XAttribute("rank", "1"),
                                    new XAttribute("percentage", "100"),
                                    new XAttribute("rsi", "0"),
                                    new XAttribute("nominalRsi", "0")))),
                        new XElement("ExteriorAddedInsulation",
                            new XAttribute("nominalInsulation", "0"),
                            new XElement("Description", "User specified"),
                            new XElement("Composite",
                                new XElement("Section",
                                    new XAttribute("rank", "1"),
                                    new XAttribute("percentage", "100"),
                                    new XAttribute("rsi", "0"),
                                    new XAttribute("nominalRsi", "0"))))),
                    new XElement("Measurements",
                        new XAttribute("height", WallHeight),
                        new XAttribute("depth", WallDepth),
                        new XAttribute("ponyWallHeight", PonyWallHeight))));
        }

        public void AddPonyWall(XElement bsmt)
        {
            if (bsmt != null)
            {
                bsmt.Element("Wall").Element("Construction").Add(
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
    }
}
