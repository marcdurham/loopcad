using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Linq;

namespace LoopCAD.WPF
{
    public class FloorTagBuilder
    {
        public static void Insert(Point3d point)
        {
            var boxes = ElevationBox.InsideElevationBoxes(point);
            if (boxes.Count == 0)
            {
                Editor().WriteMessage("\nError!  You must insert a floor tag inside an elevation box.");
                return;
            }

            PromptStringOptions floorNameOptions = new PromptStringOptions("Enter floor name")
            {
                DefaultValue = "Main Floor",
                AllowSpaces = true,
            };

            PromptResult floorNameResult = Editor()
                .GetString(floorNameOptions);
            
            FloorTag.Insert(point, floorNameResult.StringResult, boxes.First().Elevation);

            // TODO: Maybe check to see if there is already a floor tag in this
            // elevation box?
        }
        
        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
