using System;
using Autodesk.Revit.DB;

namespace ToolsV2Classes
{
    /// <summary>
    /// An integer-based 2D point class.
    /// </summary>
    class Point2dInt : IComparable<Point2dInt>
    {
        public int X { get; set; }
        public int Y { get; set; }

        const double _feet_to_mm = 25.4 * 12;

        static int ConvertFeetToMillimetres(double d)
        {
            return (int)(_feet_to_mm * d + 0.5);
        }

        /// <summary>
        /// Convert a 3D Revit XYZ to a 2D millimetre 
        /// integer point by discarding the Z coordinate
        /// and scaling from feet to mm.
        /// </summary>
        public Point2dInt(XYZ p)
        {
            X = ConvertFeetToMillimetres(p.X);
            Y = ConvertFeetToMillimetres(p.Y);
        }

        public int CompareTo(Point2dInt a)
        {
            int d = X - a.X;

            if (0 == d)
            {
                d = Y - a.Y;
            }
            return d;
        }

        public override string ToString()
        {
            return string.Format("({0},{1})", X, Y);
        }
    }
}
