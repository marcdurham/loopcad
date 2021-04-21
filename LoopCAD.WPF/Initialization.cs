using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

namespace LoopCAD.WPF
{
    public class Initialization : IExtensionApplication
    {
        public void Initialize()
        {
           
        }

        public void Terminate()
        {
            
        }

        [CommandMethod("LABEL-NODES")]
        public void LabelNodesCommand()
        {
            Editor().WriteMessage("\nLabeling nodes...");

            Transaction trans = ModelSpace.StartTransaction();

            var headLabeler = new Labeler(trans, "HEADNUMBER", "HeadLabel2", "HeadLabels", ColorIndices.Magenta);
            var teeLabeler = new Labeler(trans, "TEENUMBER", "TeeLabel2", "TeeLabels", ColorIndices.Green);

            int headNumber = 1;
            int teeNumber = 1;
            int riserNumber = 1;
            foreach (var objectId in ModelSpace.From(trans))
            {
                if (IsHead(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    headLabeler.CreateLabel($"H.{headNumber++}", block.Position);
                }
                else if (IsTee(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    teeLabeler.CreateLabel($"T.{teeNumber++}", block.Position);
                }
                else if (IsRiser(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    teeLabeler.CreateLabel($"R.{riserNumber++}", block.Position); // TODO: Add Suffix!
                }
            }

            Editor().WriteMessage($"\n{headNumber} heads labeled.");
            Editor().WriteMessage($"\n{teeNumber} tees labeled.");
            trans.Commit();
        }

        [CommandMethod("LABEL-PIPES")]
        public void LabelPipesCommand()
        {
            Editor().WriteMessage("\nLabeling pipes...");

            Transaction trans = ModelSpace.StartTransaction();

            var pipeLabeler = new Labeler(trans, "PIPENUMBER", "PipeLabel2", "PipeLabels", ColorIndices.Blue)
            {
                TextHeight = 3.0,
                XOffset = 0.0,
                YOffset = 0.0,
                HorizontalMode = TextHorizontalMode.TextCenter
            };

            int pipeNumber = 1;
            foreach (var objectId in ModelSpace.From(trans))
            {
                if (IsPipe(trans, objectId))
                {
                    var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                    for (int i = 1; i < polyline.NumberOfVertices; i++)
                    {
                        Point3d lastVertex = polyline.GetPoint3dAt(i - 1);
                        Point3d vertex = polyline.GetPoint3dAt(i);
                        pipeLabeler.CreateLabel(
                            text: $"P{pipeNumber}",
                            position: Midpoint(vertex, lastVertex));
                    }

                    pipeNumber++;
                }
            }

            trans.Commit();
            Editor().WriteMessage($"\n{pipeNumber} pipes labeled.");
        }

        [CommandMethod("RISER")]
        public void InsertRiserCommand()
        {
            Editor().WriteMessage("\nInserting riser...");

            var options = new PromptPointOptions(
                "Click location to insert riser")
            {
                AllowArbitraryInput = true,
            };

            var point = Editor().GetPoint(options);

            if (point.Status == PromptStatus.OK)
            {
                Riser.Insert(point.Value);
            }
        }

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }

        bool IsHead(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" && 
                string.Equals(block.Layer, "Heads", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("HEAD");
        }

        bool IsTee(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Tees", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("TEE");
        }

        bool IsRiser(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                (string.Equals(block.Layer, "Risers", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(block.Layer, "FloorConnectors", StringComparison.OrdinalIgnoreCase))&&
                (block.Name.ToUpper().StartsWith("FLOORCONNECTOR") || block.Name.ToUpper().StartsWith("RISER"));
        }

        bool IsPipe(Transaction trans, ObjectId objectId)
        {
            var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;
            return objectId.ObjectClass.DxfName == "LWPOLYLINE" &&
                string.Equals(polyline.Layer, "Pipes", StringComparison.OrdinalIgnoreCase);
        }

        Point3d Midpoint(Point3d a, Point3d b)
        {
            return new Point3d(
                x: (b.X - a.X) / 2 + a.X,
                y: (b.Y - a.Y) / 2 + a.Y,
                z: 0);
        }
    }
}
