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

            var headLabeler = new Labeler("HEADNUMBER", "HeadLabel2", "HeadLabels", ColorIndices.Magenta);
            var teeLabeler = new Labeler("TEENUMBER", "TeeLabel2", "TeeLabels", ColorIndices.Green);

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

            var pipeLabeler = new Labeler("PIPENUMBER", "PipeLabel2", "PipeLabels", ColorIndices.Blue)
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
                //Point3d lastVertex;
                //Point3d vertex;
                var vertices = new List<Point3d>();
                using (var trans = ModelSpace.StartTransaction())
                {
                    if (IsPipe(trans, objectId))
                    {
                        var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                        //for (int i = 1; i < polyline.NumberOfVertices; i++)
                        for (int i = 0; i < polyline.NumberOfVertices; i++)
                        {
                            vertices.Add(polyline.GetPoint3dAt(i));
                            ////lastVertex = polyline.GetPoint3dAt(i - 1);
                            ////vertex = polyline.GetPoint3dAt(i);
                            //pipeLabeler.CreateLabel(
                            //    text: $"P{pipeNumber}",
                            //    position: Midpoint(vertex, lastVertex));
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

                var floorTags = FloorTag.GetFloorTags();
                if(floorTags.Count > 1)
                {
                    // Pick which floors
                    //var dialog = new ChooseFloor(floorTags);
                    //bool? result = dialog.ShowDialog();
                    var pko = new PromptKeywordOptions("");
                    // Add our keywords
                    foreach (var ft in floorTags)
                    {
                        pko.Keywords.Add(ft.Name);
                    }

                    //// Set our prompts to include our keywords
                    ////string kws = pso.Keywords.GetDisplayString(true);

                    //pso.KeywordInput +=

                    //   delegate (object sender, SelectionTextInputEventArgs e)

                    //   {

                    //       Editor().WriteMessage("\nKeyword entered: {0}", e.Input);

                    //   };

                    //PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                    pko.Message = "\nPick a floor ";
                    //pKeyOpts.Keywords.Add("Line");
                    //pKeyOpts.Keywords.Add("Circle");
                    //pKeyOpts.Keywords.Add("Arc");
                    //pKeyOpts.AllowNone = false;

                    PromptResult result = Editor().GetKeywords(pko);

                    FloorTag floor;
                    if(result.Status == PromptStatus.OK)
                    {
                        floor = floorTags
                            .SingleOrDefault(
                                t => string.Equals(
                                    t.Name, 
                                    result.StringResult, 
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    //PromptSelectionResult result = Editor().GetSelection(pso);
                    //string selectedFloor = string.Empty;
                    //foreach(var val in result.Value)
                    //{
                    //    selectedFloor = val as string;
                    //}

                    //////var floor = floorTags
                    //////    .SingleOrDefault(t => string.Equals(t.Name, selectedFloor, StringComparison.OrdinalIgnoreCase));
                }

                foreach (var ft in floorTags)
                {
                    var newPoint = new Point3d(
                        x: ft.Position.X + offset.X,
                        y: ft.Position.Y + offset.Y,
                        0);

                    Riser.Insert(newPoint);

                    int number = RiserLabel.HighestNumber() + 1;
                    new Labeler(
                            "RISERNUMBER",
                            RiserLabel.BlockName,
                            RiserLabel.LayerName,
                            ColorIndices.Cyan)
                        .CreateLabel($"R.{number}.X", position: newPoint);
                }
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
