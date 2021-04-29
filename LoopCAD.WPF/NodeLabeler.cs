using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class NodeLabeler
    {
        public static void Run()
        {
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
                }

                //Editor().WriteMessage($"\n{headNumber} heads labeled.");
                //Editor().WriteMessage($"\n{teeNumber} tees labeled.");
                trans.Commit();
            }
        }

        static bool IsHead(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Heads", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("HEAD");
        }

        static bool IsTee(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Tees", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("TEE");
        }
    }
}
