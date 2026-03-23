using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HotPort.Models
{
    internal class BuilderValues
    {
        public string? WallRValue { get; private set; }
        public string? FloorRValue { get; private set; }
        public string? CeilingRValue { get; private set; }
        public string? VaultRValue { get; private set; }
        public string? CathedralRValue { get; private set; }
        public string? FlatCeilingRValue { get; private set; }
        public string? TallWallRValue { get; private set; }

        public BuilderValues() { }


    }
}
