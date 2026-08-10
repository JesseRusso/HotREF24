using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Navigation;
using System.Xml.Linq;

namespace HotPort
{
    internal class Door
    {
        private string _name;
        public string Name { get { return _name; } private set { _name = value; } }
        private int _width;
        public int Width { get { return _width; } private set { _width = value; } }
        private int _height;
        public int Height { get { return _height; } private set { _height = value; } }

        private int _id;
        public int Id { get { return _id; } private set { _id = value; } }


        public Door(string name, int width, int height, int id)
        {
            _height = height;
            _name = name;
            _width = width;
            _id = id;
        }
        public static XElement FrontDoor(double width, double height, int id)
        {
            XElement door = new XElement("Door",
                                new XAttribute("rValue", "0.9809"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", id),
                                new XElement("Label", "Front"),
                                new XElement("Construction",
                                    new XAttribute("energyStar", "false"),
                                    new XElement("Type",
                                        new XAttribute("code", "3"),
                                        new XAttribute("value", "0.9809"),
                                        new XElement("English", "Steel fibreglass core"),
                                        new XElement("French", "Acier / âme en fibre de verre"))),
                                new XElement("Measurements",
                                    new XAttribute("height", height),
                                    new XAttribute("width", width)));

            return door;
        }

        public static XElement PolystyreneDoor(double width, double height, string label, int id)
        {
            XElement door = new XElement("Door",
                    new XAttribute("rValue", "0.9809"),
                    new XAttribute("adjacentEnclosedSpace", "false"),
                    new XAttribute("id", id),
                    new XElement("Label", label),
                    new XElement("Construction",
                        new XAttribute("energyStar", "false"),
                        new XElement("Type",
                            new XAttribute("code", "4"),
                            new XAttribute("value", "0.9809"),
                            new XElement("English", "Steel polystyrene core"),
                            new XElement("French", "Acier / âme en polystyrène"))),
                    new XElement("Measurements",
                        new XAttribute("height", height),
                        new XAttribute("width", width)));

            return door;
        }

        public static void AddTransom(XDocument house, XElement door, int id)
        {
            XElement? comp = door.Element("Components");
            if (comp == null)
            {
                door.Add(new XElement("Components"));
                comp = door.Element("Components");
            }
            bool wide = Double.TryParse(door.Element("Measurements").Attribute("width").Value, out double width);

            XElement windowBlock = new XElement("Window",
                new XAttribute("number", "1"),
                new XAttribute("er", "28.0884"),
                new XAttribute("shgc", "0.4863"),
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", id),
                    new XElement("Label", "Transom"),
                    new XElement("Construction",
                    new XAttribute("energyStar", "false"),
                        new XElement("Type", "P2EA",
                        new XAttribute("idref", $"Code {AddDoorWindowCode(house)}"),
                        new XAttribute("rValue", "0.4019"))),
                    new XElement("Measurements",
                    new XAttribute("height", "304.8"),
                    new XAttribute("width", width * 1000),
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
                    new XAttribute("code", "5"),
                    new XElement("English", "North"),
                    new XElement("French", "Nord")));
            comp.Add(windowBlock);
        }
        private static string AddDoorWindowCode(XDocument house)
        {
            XElement? codesEl = house.Root?.Element("Codes");
            XElement? windowCodes = house.Root?.Element("Codes")?.Element("Window");
            XElement? favCodes = house.Root.Element("Codes")?.Element("Window")?.Element("Favorite");
            string codeId;

            if(codesEl == null)
            {
                house.Root.Add(new XElement("Codes"));
            }
            if(windowCodes == null)
            {
                house.Root.Element("Codes").Add(new XElement("Window"));
            }
            if (favCodes == null || favCodes == default)
            {
                house.Root.Element("Codes").Element("Window").Add(new XElement("Favorite"));
            }
            XElement? p2ea = favCodes?.Descendants("Code")
                .Where(e => e.Element("Label").Value.Equals("P2EA")).FirstOrDefault();
            if(p2ea != null || p2ea != default)
            {
                codeId = p2ea.Attribute("id").ToString().Split(" ").Last();
            }
            else
            {
                XElement code = CodeTools.DoorWindowCode(house);
                codeId = code.Attribute("id").Value.ToString().Split(" ").Last();
                house.Root.Element("Codes").Element("Window").Element("Favorite").Add(code);
            }
            return codeId;
        }
    }
}
