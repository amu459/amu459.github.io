using System;
using System.Collections.Generic;

namespace ToolsV2Classes
{
    class JtLoop : List<Point2dInt>
    {
        public JtLoop(int capacity)
          : base(capacity)
        {
        }

        public override string ToString()
        {
            return string.Join(", ", this);
        }

        //public Add( XYZ p )
        //{
        //}
    }

    class JtLoops : List<JtLoop>
    {
        public JtLoops(int capacity)
          : base(capacity)
        {
        }
    }
}
