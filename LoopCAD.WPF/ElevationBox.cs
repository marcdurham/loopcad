using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;

namespace LoopCAD.WPF
{
    public class ElevationBox
    {
        public const string LayerName = "Elevation Boxes";
        public const string BlockName = "ElevationBox";

        public static List<ObjectId> GetElevationBoxes(Transaction transaction)
        {
            var labels = new List<ObjectId>();
            foreach (var objectId in ModelSpace.From(transaction))
            {
                if (IsElevatioNBox(transaction, objectId))
                {
                    labels.Add(objectId);
                }
            }

            return labels;
        }

        static bool IsElevatioNBox(Transaction transaction, ObjectId objectId)
        {
            var block = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "POLYLINE" &&
                string.Equals(block.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(block.Name, BlockName, StringComparison.OrdinalIgnoreCase);
        }
    }
}
