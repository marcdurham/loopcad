using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class TeeLabeler
    {
        public static int Run()
        {
            var labeler = new Labeler("TEENUMBER", "TeeLabel", "TeeLabels", ColorIndices.Green)
            {
                TextHeight = 4.8
            };

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

                int teeNumber = 1;
                foreach (var objectId in ModelSpace.From(trans))
                {
                    if (IsTee(trans, objectId))
                    {
                        var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                        labeler.CreateLabel($"T.{teeNumber++}", block.Position);
                    }
                }

                trans.Commit();

                return teeNumber;
            }
        }

        static bool IsTee(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Tees", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("TEE");
        }

        static bool IsLabel(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "TeeLabels", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(block.Name, "TeeLabel", StringComparison.OrdinalIgnoreCase);
        }
    }
}
