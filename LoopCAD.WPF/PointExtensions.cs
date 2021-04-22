using Autodesk.AutoCAD.Geometry;
using System;

namespace LoopCAD.WPF
{
    public class PointExtensions
    {
        public static bool IsNear(Point3d a, Point3d b, double distance = 5.0)
        {
            return Math.Abs(a.X - b.X) <= distance &&
                Math.Abs(a.Y - b.Y) <= distance;
        }
    }
}
