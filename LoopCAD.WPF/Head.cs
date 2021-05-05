using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class Head
    {
        public const string BlockName = "HeadTest16";
        public const string Layer = "Heads";
        readonly Database db;
        readonly BlockTable table;
        readonly Transaction transaction;

        public int Coverage { get; set; }

        public Head(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            table = transaction.GetObject(
                db.BlockTableId,
                OpenMode.ForWrite) as BlockTable;
        }

        public void InsertAt(Point3d position, string model, int coverage)
        {
            Coverage = Coverage;

            BlockTableRecord record = Define();

            var blockRef = new BlockReference(position, Define().Id)
            {
                Layer = Layer,
                ColorIndex = ColorIndices.ByLayer
            };

            ModelSpace.From(transaction).AppendEntity(blockRef);
            transaction.AddNewlyCreatedDBObject(blockRef, true);

            foreach (ObjectId id in record)
            {
                using (var def = id.GetObject(OpenMode.ForRead) as AttributeDefinition)
                {
                    if ((def != null) && (!def.Constant)
                        && def.Tag.ToUpper() == "MODEL")
                    {
                        using (var ar = new AttributeReference())
                        {
                            ar.SetAttributeFromBlock(def, blockRef.BlockTransform);
                            ar.TextString = $"{model}-{coverage}";

                            blockRef.AttributeCollection.AppendAttribute(ar);
                            transaction.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                }
            }

           // transaction.Commit();
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
                record = transaction.GetObject(
                    table[BlockName], 
                    OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }

        BlockTableRecord DefinitionFrom(BlockTable table)
        {
            WPF.Layer.Ensure(Layer, ColorIndices.Red);
            WPF.Layer.Ensure("HeadCoverage", ColorIndices.Yellow);

            var record = new BlockTableRecord
            {
                Name = BlockName
            };

            var circle = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 6.6, // inches
                Layer = Layer,
                ColorIndex = ColorIndices.ByLayer,
            };

            record.AppendEntity(circle);

            var inner = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 2.3, // inches
                Layer = Layer,
                ColorIndex = ColorIndices.ByLayer,
            };

            record.AppendEntity(inner);

            var attribute = new AttributeDefinition
            {
                Tag = "MODEL",
                TextString = "MODEL-123",
                Position = new Point3d(10, 4, 0),
                Height = 8.0,
                ColorIndex = ColorIndices.ByLayer,
                Layer = Layer
            };

            record.AppendEntity(attribute);

            // TODO: Add square
            var square = new Polyline(4)
            {
                Layer = "HeadCoverage",
                Closed = true,
                ColorIndex = ColorIndices.Yellow,
            };

            double radius = 6 * 12;

            square.AddVertexAt(0, new Point2d(-radius, radius), 0, 0, 0);
            square.AddVertexAt(1, new Point2d(radius, radius), 0, 0, 0);
            square.AddVertexAt(2, new Point2d(radius, -radius), 0, 0, 0);
            square.AddVertexAt(3, new Point2d(-radius, -radius), 0, 0, 0);

            record.AppendEntity(square);

            // TODO: Add coverage label

            table.Add(record);

            return record;
        }
    }
}
