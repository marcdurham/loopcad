using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class Head
    {
        public const string BlockName = "Head";
        public const string Layer = "Heads";
        public const string CoverageLayer = "HeadCoverage";
        readonly Database db;
        readonly BlockTable table;
        readonly Transaction transaction;

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
            BlockTableRecord record = Define(coverage);

            var blockRef = new BlockReference(position, Define(coverage).Id)
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

        public BlockTableRecord Define(int coverage)
        {
            BlockTableRecord record;

            if (!table.Has($"{BlockName}{coverage}"))
            {
                record = DefinitionFrom(table, coverage);
                transaction.AddNewlyCreatedDBObject(record, true);
            }
            else
            {
                record = transaction.GetObject(
                    table[$"{BlockName}{coverage}"], 
                    OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }

        BlockTableRecord DefinitionFrom(BlockTable table, int coverage)
        {
            WPF.Layer.Ensure(Layer, ColorIndices.Red);
            WPF.Layer.Ensure(CoverageLayer, ColorIndices.Yellow);

            var record = new BlockTableRecord
            {
                Name = $"{BlockName}{coverage}"
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
                Layer = CoverageLayer,
                Closed = true,
                ColorIndex = ColorIndices.ByLayer,
            };

            double radius = (coverage/2) * 12; // convert to inches

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
