using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;
using System.Collections.Generic;

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

            var headLabeler = new Labeler("HEADNUMBER", "HeadLabel", "HeadLabels", ColorIndices.Blue)
            {
                TextHeight = 4.8
            };

            var teeLabeler = new Labeler("TEENUMBER", "TeeLabel", "TeeLabels", ColorIndices.Green)
            {
                TextHeight = 4.8
            };

            using (var trans = ModelSpace.StartTransaction())
            {
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
        }

        [CommandMethod("LABEL-PIPES")]
        public void LabelPipesCommand()
        {
            Editor().WriteMessage("\nLabeling pipes...");

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
                foreach(ObjectId id in ModelSpace.From(trans))
                {
                    ids.Add(id);
                }
            }

            foreach (ObjectId objectId in ids)
            {
                var vertices = new List<Point3d>();
                using (var trans = ModelSpace.StartTransaction())
                {
                    if (IsPipe(trans, objectId))
                    {
                        var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                        for (int i = 0; i < polyline.NumberOfVertices; i++)
                        {
                            vertices.Add(polyline.GetPoint3dAt(i));
                        }

                        pipeNumber++;
                    }

                    trans.Commit();
                }

                for (int i = 1; i < vertices.Count; i++)
                {
                    pipeLabeler.CreateLabel(
                        text: $"P{pipeNumber}",
                        position: Midpoint(vertices[i], vertices[i-1]));
                }
            }

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
                var boxes = ElevationBox.InsideElevationBoxes(point.Value);
                if (boxes.Count == 0)
                {
                    Editor().WriteMessage("\nError!  You must insert a riser inside an elevation box.");
                    return;
                }

                if (boxes.Count(b => b.FloorTag != null) != 1)
                {
                    Editor().WriteMessage("\nError!  You must insert a floor tag on each floor.");
                    return;
                }

                var floorTag = boxes.Select(b => b.FloorTag).Single();
                var offset = new Point3d(
                    x: point.Value.X - floorTag.Position.X,
                    y: point.Value.Y - floorTag.Position.Y,
                    0);

                var allFloorTags = FloorTag.GetFloorTags();
                if(allFloorTags.Count > 10)
                {
                    Editor().WriteMessage("\nError! LoopCAD cannot handle more than ten floors.");
                    return;
                }

                var selectedFloorTags = new List<FloorTag>()
                {
                    floorTag
                };

                if(allFloorTags.Count > 1)
                {
                    var pko = new PromptKeywordOptions("");
                    foreach (var ft in allFloorTags)
                    {
                        if (ft.Name != floorTag.Name)
                        {
                            pko.Keywords.Add($"{allFloorTags.IndexOf(ft)}: {ft.Name}");
                        }
                    }

                    pko.Message = "\nPick a floor ";
                    pko.AllowNone = false;

                    PromptResult result = Editor().GetKeywords(pko);

                    if(result.Status == PromptStatus.OK)
                    {
                        int floorIndex = int.Parse(result.StringResult[0].ToString());
                        selectedFloorTags.Add(allFloorTags[floorIndex]);
                    }
                }

                if(selectedFloorTags.Count != 2)
                {
                    Editor().WriteMessage("Error! The riser could not be inserted on two floors.");

                    return;
                }

                char suffix = (char)((byte)RiserLabel.HighestSuffix() + 1);
                int number = RiserLabel.HighestNumber() + 1;

                foreach (var ft in selectedFloorTags)
                {
                    // This point will be disposed, so clone it
                    var newPoint = new Point3d(
                        x: ft.Position.X + offset.X,
                        y: ft.Position.Y + offset.Y,
                        z: ft.Position.Z + offset.Z);

                    Riser.Insert(newPoint);

                    var labeler = new Labeler(
                        RiserLabel.TagName,
                        RiserLabel.BlockName,
                        RiserLabel.LayerName,
                        ColorIndices.Cyan)
                    {
                         XOffset = 15.0,
                         TextHeight = 8.0
                    };

                    labeler.CreateLabel($"R.{number}.{suffix}", position: newPoint);
                    
                    // Number increments, but suffix is shared between the
                    // upper and lower riser, and used to match them.
                    number++;
                }
            }
        }

        [CommandMethod("SAVE-DXF")]
        public void SaveAsDxfCommand()
        {
            Editor().WriteMessage("\nSaving as DXF...");

            string path = Application.DocumentManager.MdiActiveDocument.Name.Replace(".dwg", ".dxf");
            string directory = System.IO.Path.GetDirectoryName(path);
            string fileName = System.IO.Path.GetFileName(path);

            var dialog = new Microsoft.Win32.SaveFileDialog()
            {
                Filter = "Drawing EXchange Files (*.dxf)|*.dxf|All Files (*.*)|*.*",
                DefaultExt = ".dxf",
                FileName = fileName,
                InitialDirectory = directory
            };

            var result = dialog.ShowDialog();
            if (result.HasValue)
            {
                HostApplicationServices.WorkingDatabase.DxfOut(fileName: dialog.FileName, precision: 8, DwgVersion.Current);
                Editor().WriteMessage("\nDone");
            }
            else
            {
                Editor().WriteMessage("\nSave As DXF cancelled");
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
