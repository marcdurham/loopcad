using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace LoopCAD.WPF
{
    public class FloorTag
    {
        public const string LayerName = "Floor Tags";
        public const string BlockName = "FloorTag";

        public Point3d Position { get; set; }
        public string Name { get; set; }
        public int Elevation { get; set; }
        public List<ElevationBox> ElevationBoxes { get; set; }
            = new List<ElevationBox>();

        public static List<FloorTag> GetFloorTags()
        {
            using (var transaction = ModelSpace.StartTransaction())
            {
                var floorTags = new List<FloorTag>();
                foreach (ObjectId id in GetFloorTagIds(transaction))
                {
                    var block = transaction.GetObject(id, OpenMode.ForRead) as BlockReference;
                    var elevationString = AttributeReader.TextString(
                        transaction,
                        id,
                        BlockName,
                        tag: "ELEVATION");

                    int.TryParse(elevationString, out int elevation);
                    floorTags.Add(new FloorTag
                    {
                        Position = new Point3d(
                            x: block.Position.X,
                            y: block.Position.Y,
                            z: block.Position.Z),
                        Name = AttributeReader.TextString(
                            transaction,
                            id,
                            BlockName,
                            tag: "NAME"),

                        Elevation = elevation
                    });
                }

                floorTags = floorTags.OrderBy(t => t.Elevation).ToList();

                return floorTags;
            }
        }

        public override string ToString()
        {
            return Name;
        }

        static List<ObjectId> GetFloorTagIds(Transaction transaction)
        {
            var floorTags = new List<ObjectId>();
            foreach (var objectId in ModelSpace.From(transaction))
            {
                if (IsFloorTagLabel(transaction, objectId))
                {
                    floorTags.Add(objectId);
                }
            }

            return floorTags;
        }

        static bool IsFloorTagLabel(Transaction transaction, ObjectId objectId)
        {
            var block = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                (string.Equals(block.Layer, LayerName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(block.Layer, "Floor Tags", StringComparison.OrdinalIgnoreCase)) &&
                string.Equals(block.Name, BlockName, StringComparison.OrdinalIgnoreCase);
        }
    }
}
