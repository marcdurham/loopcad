using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

namespace LoopCAD.WPF
{
    public class PipeLabeler
    {
        public static int Run()
        {
            var pipeLabeler = new Labeler("PIPENUMBER", "PipeLabel", "PipeLabels", ColorIndices.Blue)
            {
                TextHeight = 3.0,
                XOffset = 0.0,
                YOffset = 0.0,
                HorizontalMode = TextHorizontalMode.TextCenter
            };

            int pipeNumber = 1;

            var ids = new List<ObjectId>();
            using (var trans = ModelSpace.StartTransaction())
            {
                foreach (ObjectId id in ModelSpace.From(trans))
                {
                    ids.Add(id);
                }
            }

            using (var trans = ModelSpace.StartTransaction())
            {
                foreach (ObjectId objectId in ids)
                {
                    if (IsLabel(trans, objectId))
                    {
                        var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                        block.Erase(true);
                    }
                }

                foreach (ObjectId objectId in ids)
                {
                    var vertices = new List<Point3d>();
                    if (IsPipe(trans, objectId))
                    {
                        var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                        for (int i = 0; i < polyline.NumberOfVertices; i++)
                        {
                            vertices.Add(polyline.GetPoint3dAt(i));
                        }

                        pipeNumber++;
                    }

                    for (int i = 1; i < vertices.Count; i++)
                    {
                        pipeLabeler.CreateLabel(
                            text: $"P{pipeNumber}",
                            position: Midpoint(vertices[i], vertices[i - 1]));
                    }
                }

                trans.Commit();

                return pipeNumber;
            }
        }

        static bool IsPipe(Transaction trans, ObjectId objectId)
        {
            var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;
            return objectId.ObjectClass.DxfName == "LWPOLYLINE" &&
                string.Equals(polyline.Layer, "Pipes", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsLabel(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (block != null)
            {
                return string.Equals(block.Layer, "PipeLabels", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(block.Name, "PipeLabel", StringComparison.OrdinalIgnoreCase);
            }

            var text = trans.GetObject(objectId, OpenMode.ForRead) as DBText;
            if (text != null)
            {
                return string.Equals(block.Layer, "Pipe Labels", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        static Point3d Midpoint(Point3d a, Point3d b)
        {
            return new Point3d(
                x: (b.X - a.X) / 2 + a.X,
                y: (b.Y - a.Y) / 2 + a.Y,
                z: 0);
        }
    }
}
