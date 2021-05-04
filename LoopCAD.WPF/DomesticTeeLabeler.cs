using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class DomesticTeeLabeler
    {
        public static int Run()
        {
            var labeler = new Labeler(
                new LabelSpecs
                {
                    Tag = "TEENUMBER",
                    BlockName = "DomesticTeeLabel",
                    Layer = "DomesticTeeLabels",
                    LayerColorIndex = ColorIndices.Green
                });

            using (var trans = ModelSpace.StartTransaction())
            {
                foreach (var objectId in ModelSpace.From(trans))
                {
                    if (IsLabel(trans, objectId))
                    {
                        var block = trans.GetObject(objectId, OpenMode.ForWrite) as BlockReference;

                        block.Erase(true);
                    }
                }

                int number = 1;
                foreach (var objectId in ModelSpace.From(trans))
                {
                    if (IsTee(trans, objectId))
                    {
                        var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                        labeler.CreateLabel($"D.T.{number++}", block.Position);
                    }
                }

                trans.Commit();

                return number;
            }
        }

        static bool IsTee(Transaction trans, ObjectId objectId)
        {
            if (objectId.IsErased)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (block == null)
            {
                return false;
            }

            return
                string.Equals(block.Layer, "DomesticTees", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("DOMESTICTEE");
        }

        static bool IsLabel(Transaction trans, ObjectId objectId)
        {
            if (objectId.IsErased)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (block == null)
            {
                return false;
            }

            return
                string.Equals(block.Layer, "DomesticTeeLabels", StringComparison.OrdinalIgnoreCase);
        }
    }
}
