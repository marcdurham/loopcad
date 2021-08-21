using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace LoopCAD.WPF
{
    public class PipeLabeler
    {
        public static int LabelAllPipes()
        {
            var pipeLabeler = new Labeler(
                 new LabelSpecs
                 {
                     Tag = "PIPENUMBER",
                     BlockName = "PipeLabel",
                     Layer = "PipeLabels",
                     LayerColorIndex = ColorIndices.Blue,
                     TextHeight = 4.0,
                     XOffset = 0.0,
                     YOffset = 0.0,
                     HorizontalMode = TextHorizontalMode.TextCenter
                 });

            int pipeNumber = 1;
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("  LabelAllPipes starting...");
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
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("  Finding labels...");
                foreach (ObjectId objectId in ids)
                {
                    if (IsLabel(trans, objectId))
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"    label found, erase block...");
                        var block = trans.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                        var blkrefClass = RXObject.GetClass(typeof(BlockReference));
                        if (block != null && objectId.ObjectClass == blkrefClass)
                        {
                            block.Erase(true);
                        }

                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"    label found, erase DBTtext...");
                        var text = trans.GetObject(objectId, OpenMode.ForWrite) as DBText;
                        if (text != null)
                        {
                            text.Erase(true);
                        }
                    }
                }

                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("  Finding pipes...");
                foreach (ObjectId objectId in ids)
                {
                    var vertices = new List<Point3d>();
                    if (IsPipe(trans, objectId))
                    {
                        var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"    Adding {polyline.NumberOfVertices} vertices...");
                        for (int i = 0; i < polyline.NumberOfVertices; i++)
                        {
                            vertices.Add(polyline.GetPoint3dAt(i));
                        }

                        pipeNumber++;
                    }

                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"    Creating label: p{pipeNumber}...");
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
            if(objectId.IsNull || objectId.IsErased || !objectId.IsValid)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            var blkrefClass = RXObject.GetClass(typeof(BlockReference));
            if (block != null && objectId.ObjectClass == blkrefClass)
            {
                return string.Equals(block.Layer, "PipeLabels", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(block.Name, "PipeLabel", StringComparison.OrdinalIgnoreCase);
            }

            var text = trans.GetObject(objectId, OpenMode.ForRead) as DBText;
            if (text != null)
            {
                return string.Equals(text.Layer, "Pipe Labels", StringComparison.OrdinalIgnoreCase);
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
