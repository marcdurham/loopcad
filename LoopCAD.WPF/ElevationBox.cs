using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LoopCAD.WPF
{
    public class ElevationBox
    {
        public const string LayerName = "ElevationBox";

        public ElevationBox(Polyline box)
        {
            Id = box.Id;
            SetSides(box);
        }

        public ObjectId Id { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public double Left { get; set; }
        public double Right { get; set; }
        public int Elevation { get; set; }

        public static List<ElevationBox> InsideElevationBoxes(Point3d point)
        {
            using (var transaction = ModelSpace.StartTransaction())
            {
                var boxes = GetElevationBoxes(transaction);
                var inside = new List<ElevationBox>();
                foreach (var box in boxes)
                {
                    if (box.IsInside(point))
                    {
                        inside.Add(box);
                    }
                }

                return inside;
            }
        }

        public double Area()
        {
            return Math.Abs(Top - Bottom) * Math.Abs(Right - Bottom);
        }

        static List<ElevationBox> GetElevationBoxes(Transaction transaction)
        {
            var boxes = new List<ObjectId>();
            var labels = new List<ObjectId>();
            foreach (var objectId in ModelSpace.From(transaction))
            {
                if (IsElevationBoxPolyline(transaction, objectId))
                {
                    boxes.Add(objectId);
                }
                else if(IsElevationBoxLabel(transaction, objectId))
                {
                    labels.Add(objectId);
                }
            }

            var elevationsBoxes = new List<ElevationBox>();
            foreach(var box in boxes)
            {
                var polyline = transaction.GetObject(box, OpenMode.ForRead) as Polyline;
                foreach (var label in labels)
                {
                    var mtext = transaction.GetObject(label, OpenMode.ForRead) as MText;
                    if(mtext == null)
                    {
                        continue;
                    }

                    for (int i = 0; i < polyline.NumberOfVertices; i++)
                    {
                        if(PointExtensions.IsNear(mtext.Location, polyline.GetPoint3dAt(i))
                            && !elevationsBoxes.Exists(b => b.Id == polyline.Id))
                        {
                            var match = Regex.Match(
                                mtext.Text,
                                @"Elevation (\d+)",
                                RegexOptions.IgnoreCase);

                            if(!match.Success)
                            {
                                continue;
                            }

                            int elevation = int.Parse(match.Groups[1].Value);

                            var b = new ElevationBox(polyline)
                            {
                                Id = polyline.Id,
                                Elevation = elevation
                            };

                            b.SetSides(polyline);
                            elevationsBoxes.Add(b);

                            continue;
                        }
                    }
                }
            }

            return elevationsBoxes;
        }

        public bool IsInside(Point3d point)
        {
            return point.X >= Left &&
                point.X <= Right &&
                point.Y >= Bottom &&
                point.Y <= Top;
        }

        static bool IsElevationBoxPolyline(Transaction transaction, ObjectId objectId)
        {
            var polyline = transaction.GetObject(objectId, OpenMode.ForRead) as Polyline;
            return (objectId.ObjectClass.DxfName == "POLYLINE" || 
                    objectId.ObjectClass.DxfName == "LWPOLYLINE") &&
                string.Equals(polyline.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                polyline.NumberOfVertices >= 4;
        }

        static bool IsElevationBoxLabel(Transaction transaction, ObjectId objectId)
        {
            var mtext = transaction.GetObject(objectId, OpenMode.ForRead) as MText;
            return objectId.ObjectClass.DxfName == "MTEXT" &&
                string.Equals(mtext.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                mtext.Contents != null &&
                Regex.IsMatch(mtext.Contents, @"Elevation \d+");
        }

        void SetSides(Polyline box)
        {
            if (box == null)
            {
                throw new ArgumentNullException(nameof(box));
            }

            var points = new List<Point3d>();
            for (int i = 0; i < box.NumberOfVertices; i++)
            {
                points.Add(box.GetPoint3dAt(i));
            }

            Left = points.Min(p => p.X);
            Right = points.Max(p => p.X);
            Top = points.Max(p => p.Y);
            Bottom = points.Min(p => p.Y);
        }
    }
}
