using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class HeadLabeler
    {
        public static int Run()
        {
            var labeler = new Labeler(
                 new LabelSpecs
                 {
                     Tag = "HEADNUMBER",
                     BlockName = "HeadLabel",
                     Layer = "HeadLabels",
                     LayerColorIndex = ColorIndices.Blue
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

                trans.Commit();
            }

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
            if(objectId.IsErased)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (block == null)
            {
                return false;
            }

            return
                string.Equals(block.Layer, "Heads", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("HEAD");
        }

        static bool IsLabel(Transaction trans, ObjectId objectId)
        {
            if (objectId.IsErased)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if(block == null)
            {
                return false;
            }

            // I'd check the block name, but it can't be accessed if the entity has
            // been erased, and causes and error, even though IsErased returns false
            return
                string.Equals(block.Layer, "HeadLabels", StringComparison.OrdinalIgnoreCase);
            //&&
            //    string.Equals(block.Name, "HeadLabel", StringComparison.OrdinalIgnoreCase);
        }
    }
}
