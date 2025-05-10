using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        
        public Door()
        {

        }
    }
}
