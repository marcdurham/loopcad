using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class HeadLabeler
    {
        public static int Run()
        {
            var labeler = new Labeler("HEADNUMBER", "HeadLabel", "HeadLabels", ColorIndices.Blue)
            {
                TextHeight = 4.8
            };

            using (var trans = ModelSpace.StartTransaction())
            {
                int headNumber = 1;
                foreach (var objectId in ModelSpace.From(trans))
                {
                    if (IsHead(trans, objectId))
                    {
                        var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                        labeler.CreateLabel($"H.{headNumber++}", block.Position);
                    }
                }

                trans.Commit();

                return headNumber;
            }
        }

        static bool IsHead(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Heads", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("HEAD");
        }
    }
}
