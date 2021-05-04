using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace LoopCAD.WPF
{
    public class RiserBuilder
    {
        public static void Insert(Point3d point)
        {
            var boxes = ElevationBox.InsideElevationBoxes(point);
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
                x: point.X - floorTag.Position.X,
                y: point.Y - floorTag.Position.Y,
                0);

            var allFloorTags = FloorTag.GetFloorTags();
            if (allFloorTags.Count > 10)
            {
                Editor().WriteMessage("\nError! LoopCAD cannot handle more than ten floors.");
                return;
            }

            var selectedFloorTags = new List<FloorTag>()
            {
                floorTag
            };

            if (allFloorTags.Count > 1)
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

                if (result.Status == PromptStatus.OK)
                {
                    int floorIndex = int.Parse(result.StringResult[0].ToString());
                    selectedFloorTags.Add(allFloorTags[floorIndex]);
                }
            }

            if (selectedFloorTags.Count != 2)
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
                     new LabelSpecs
                     {
                         Tag = RiserLabel.TagName,
                         BlockName = RiserLabel.BlockName,
                         Layer = RiserLabel.LayerName,
                         LayerColorIndex = ColorIndices.Cyan,
                         TextHeight = 8.0,
                         XOffset = 15.0,
                     });

                labeler.CreateLabel($"R.{number}.{suffix}", position: newPoint);

                // Number increments, but suffix is shared between the
                // upper and lower riser, and used to match them.
                number++;
            }
        }
        
        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
