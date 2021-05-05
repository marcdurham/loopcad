using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class HeadBuilder
    {
        public static void Insert(int coverage)
        {
            // TODO: int number = HeadLabel.HighestNumber() + 1;
            //int number = 1;
            var jobData = JobData.Load();

            using (var transaction = ModelSpace.StartTransaction())
            {
                var table = (BlockTable)transaction.GetObject(
                    Editor().Document.Database.BlockTableId,
                    OpenMode.ForRead);

                var jigBlock = (BlockTableRecord)transaction.GetObject(
                    table[$"Head{coverage}"],
                    OpenMode.ForRead);

                var jig = new BlockJig();
                PromptResult res = jig.DragMe(jigBlock.ObjectId, out Point3d point);

                // This point will be disposed outside of this block, so clone it
                var pointClone = new Point3d(
                    x: point.X,
                    y: point.Y,
                    z: point.Z);

                if (res.Status == PromptStatus.OK)
                {
                    new Head(transaction).InsertAt(pointClone, model: jobData?.HeadModelDefault ?? "", coverage: coverage);
                }

                //var labeler = new Labeler(
                //        new LabelSpecs
                //        {
                //            Tag = "HEADNUMBER", //RiserLabel.TagName,
                //            BlockName = "HeadLabel", //RiserLabel.BlockName,
                //            Layer = "HeadLabels", //RiserLabel.LayerName,
                //            LayerColorIndex = ColorIndices.Red,
                //            TextHeight = 5.0,
                //            XOffset = 15.0,
                //        });

                //labeler.CreateLabel($"H.{number}", position: pointClone);

                transaction.Commit();
            }

            ////number++;
        }

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
