using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace LoopCAD.WPF
{
    public class HeadBuilder
    {
        public static void Insert(Point3d point)
        {
            //char suffix = (char)((byte)RiserLabel.HighestSuffix() + 1);
            //int number = RiserLabel.HighestNumber() + 1;
            int number = 1;

            //foreach (var ft in selectedFloorTags)
            {
                // This point will be disposed, so clone it
                var newPoint = new Point3d(
                    x: point.X + 5,
                    y: point.Y + 5,
                    z: point.Z);

                Head.Insert(newPoint, model: "ABC-1000");

                var labeler = new Labeler(
                     new LabelSpecs
                     {
                         Tag = "HEADNUMBER", //RiserLabel.TagName,
                         BlockName = "HeadLabel", //RiserLabel.BlockName,
                         Layer = "HeadLabels", //RiserLabel.LayerName,
                         LayerColorIndex = ColorIndices.Red,
                         TextHeight = 5.0,
                         XOffset = 15.0,
                     });

                labeler.CreateLabel($"H.{number}", position: newPoint);

                // Number increments, but suffix is shared between the
                // upper and lower riser, and used to match them.
                ////number++;
            }
        }
        
        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
