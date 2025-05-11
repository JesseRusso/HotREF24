using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotPort
{
    internal class Window
    {
        private string _name;
        private double _width;
        private double _height;
        private double _rsi;
        private double _shgc;
        private int _id;
        private int _codeId;
        private int _floor;
        private int _uValue;
        private double _overhang;
        public string Name { get { return _name; } }
        public double Width { get { return _width; } }
        public double Height { get { return _height; } }
        public double RSI { get { return _rsi; } }
        public double Shgc { get { return _shgc; } }
        public int Id { get { return _id;} }
        public int codeId { get { return _codeId;} set { _codeId = value; } }
        public int floor { get { return _floor; } }

        //TODO initialize all values to the appropriate metric values
        public Window(string name, int width, int height, double uValue, double shgc, int floor, double overhang, int id)
        {
            _name = name;
            _width = Math.Round(width * 25.4, 6);
            _height = Math.Round(height * 25.4, 6);
            _shgc = Math.Round(shgc, 2);
            _id = id;
            _floor = floor;
            _rsi = Math.Round((1 / uValue), 4);
            _overhang = Math.Round((overhang * 0.0254) * 12, 4);
            _uValue = (int)(uValue * 100);
        }
        public XElement getWindowBlock()
        {
            XElement windowBlock = new XElement("Window",
                new XAttribute("number", "1"),
                new XAttribute("er", "-32.1684"),
                new XAttribute("shgc", "0.26"),
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", _id),
                    new XElement("Label", _name),
                    new XElement("Construction",
                    new XAttribute("energyStar", "false"),
                        new XElement("Type", ToString(),
                        new XAttribute("idref", ("Code " + codeId)),
                        new XAttribute("rValue", _rsi))),
                    new XElement("Measurements",
                    new XAttribute("height", _height),
                    new XAttribute("width", _width),
                    new XAttribute("headerHeight", "0"),
                    new XAttribute("overhangWidth", _overhang),
                        new XElement("Tilt",
                        new XAttribute("code", "1"),
                        new XAttribute("value", "90"),
                        new XElement("English", "Vertical"),
                        new XElement("French", "Verticale"))),
                    new XElement("Shading",
                    new XAttribute("curtain", "1"),
                    new XAttribute("shutterRValue", "0")),
                    new XElement("FacingDirection",
                    new XAttribute("code", "1"),
                    new XElement("English", "North"),
                    new XElement("French", "Nord")));

            return windowBlock;
        }
        public void AddWindow(XDocument house)
        {
            XElement[]? walls = new XElement[4];

            walls[0] = house?.Root?.Element("House")?.Element("Components")?.Element("Basement");

            walls[1] = (XElement)(from el in house.Root?.Element("House")?.Element("Components").Descendants("Wall")
                                        where el.Element("Label").Value.Contains("1")
                                        select el).FirstOrDefault();
            walls[2] = (XElement)(from el in house.Root?.Element("House")?.Element("Components").Descendants("Wall")
                                  where el.Element("Label").Value.Contains("2")
                                  select el).FirstOrDefault();
            XElement? third = (XElement)(from el in house.Root?.Element("House")?.Element("Components")?.Descendants("Wall")
                                         where el.Element("Label")?.Value.Contains("3") ?? false
                                         select el)?.FirstOrDefault();
            if(third != null) walls[3] = third;
            walls[_floor].Element("Components").AddFirst(getWindowBlock());
        }
        public override string ToString()
        {
            return $"u{_uValue}shg{Math.Round(Shgc * 100,2)}";
        }

    }
}
