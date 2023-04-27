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

        readonly Database db;
        readonly BlockTable table;
        readonly Transaction transaction;

        public FloorTag()
        {
        }

        public FloorTag(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            table = transaction.GetObject(
                db.BlockTableId,
                OpenMode.ForWrite) as BlockTable;
        }

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

        public static void Insert(Point3d position, string floorName, double elevationFt)
        {
            new FloorTag(StartTransaction())
                .InsertAt(position, floorName, elevationFt);
        }

        void InsertAt(Point3d position, string floorName, double elevationFt)
        {
            BlockTableRecord record = Define();

            var blockRef = new BlockReference(position, Define().Id)
            {
                Layer = LayerName,
                ColorIndex = ColorIndices.ByLayer
            };

            ModelSpace.From(transaction).AppendEntity(blockRef);
            transaction.AddNewlyCreatedDBObject(blockRef, true);


            foreach (ObjectId id in record)
            {
                using (var def = id.GetObject(OpenMode.ForRead) as AttributeDefinition)
                {
                    if ((def != null) && (!def.Constant)
                        && def.Tag.ToUpper() == "NAME")
                    {
                        using (var ar = new AttributeReference())
                        {
                            ar.SetAttributeFromBlock(def, blockRef.BlockTransform);
                            ar.TextString = $"{floorName}";
                            ar.Rotation = 0;

                            blockRef.AttributeCollection.AppendAttribute(ar);
                            transaction.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                    else if ((def != null) && (!def.Constant)
                        && def.Tag.ToUpper() == "ELEVATION")
                    {
                        using (var ar = new AttributeReference())
                        {
                            ar.SetAttributeFromBlock(def, blockRef.BlockTransform);
                            ar.TextString = $"{elevationFt}";
                            ar.Rotation = 0;

                            blockRef.AttributeCollection.AppendAttribute(ar);
                            transaction.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                }
            }

            transaction.Commit();
        }

        static Transaction StartTransaction()
        {
            return HostApplicationServices
                .WorkingDatabase
                .TransactionManager
                .StartTransaction();
        }

        BlockTableRecord Define()
        {
            BlockTableRecord record;

            if (!table.Has(BlockName))
            {
                record = DefinitionFrom(table);
                transaction.AddNewlyCreatedDBObject(record, true);
            }
            else
            {
                record = transaction.GetObject(table[BlockName], OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }


        BlockTableRecord DefinitionFrom(BlockTable table)
        {
            WPF.Layer.Ensure(LayerName, ColorIndices.Cyan);

            var record = new BlockTableRecord
            {
                Name = BlockName
            };

            // Outer circle
            var circle = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 9.0, // inches
                Layer = LayerName,
                ColorIndex = ColorIndices.ByLayer,
            };
            record.AppendEntity(circle);

            // Vertical line
            var vertical = new Polyline(2)
            {
                Layer = LayerName,
                Closed = true,
                ColorIndex = ColorIndices.ByLayer,
            };
            vertical.AddVertexAt(0, new Point2d(0.0, 9.0), 0, 0, 0);
            vertical.AddVertexAt(1, new Point2d(0.0, -9.0), 0, 0, 0);
            record.AppendEntity(vertical);

            // Horizontal line
            var horizontal = new Polyline(2)
            {
                Layer = LayerName,
                Closed = true,
                ColorIndex = ColorIndices.ByLayer,
            };
            horizontal.AddVertexAt(0, new Point2d(9.0, 0.0), 0, 0, 0);
            horizontal.AddVertexAt(1, new Point2d(-9.0, 0.0), 0, 0, 0);
            record.AppendEntity(horizontal);

            var nameAttribute = new AttributeDefinition
            {
                Tag = "NAME",
                TextString = "Floor Name",
                Position = new Point3d(9.0, -9.0, 0),
                Height = 8.0,
                ColorIndex = ColorIndices.ByLayer,
                Layer = LayerName,
            };
            record.AppendEntity(nameAttribute);
            
            var elevationAttribute = new AttributeDefinition
            {
                Tag = "ELEVATION",
                TextString = "Elevation (ft)",
                Position = new Point3d(9.0, -18.0, 0),
                Height = 8.0,
                ColorIndex = ColorIndices.ByLayer,
                Layer = LayerName,
            };
            record.AppendEntity(elevationAttribute);


            table.Add(record);

            return record;
        }
    }
}
