using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;

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

        public int coverage { get; set; }

        public Head(Transaction transaction, int coverage)
        {
            this.transaction = transaction;
            this.coverage = coverage;

            db = HostApplicationServices.WorkingDatabase;
            table = transaction.GetObject(
                db.BlockTableId,
                OpenMode.ForWrite) as BlockTable;
        }

        public void InsertAt(Point3d position, string model, bool sideWall = false, double angle = 0.0)
        {
            BlockTableRecord record = Define(sideWall);

            var blockRef = new BlockReference(position, Define(sideWall).Id)
            {
                Layer = Layer,
                ColorIndex = ColorIndices.ByLayer,
                Rotation = angle,
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
                            ar.Rotation = 0;

                            blockRef.AttributeCollection.AppendAttribute(ar);
                            transaction.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                }
            }
        }

        public BlockTableRecord Define(bool sideWall = false)
        {
            BlockTableRecord record;
            string sw = sideWall ? "SW" : "";

            if (!table.Has($"{BlockName}{sw}{coverage}"))
            {
                record = DefinitionFrom(table, coverage, sideWall);
                transaction.AddNewlyCreatedDBObject(record, true);
            }
            else
            {
                record = transaction.GetObject(
                    table[$"{BlockName}{sw}{coverage}"], 
                    OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }

        BlockTableRecord DefinitionFrom(
            BlockTable table, 
            int coverage, 
            bool sideWall = false)
        {
            WPF.Layer.Ensure(Layer, ColorIndices.Red);
            WPF.Layer.Ensure(CoverageLayer, ColorIndices.Yellow);
            
            string sw = sideWall ? "SW" : "";
            
            var record = new BlockTableRecord
            {
                Name = $"{BlockName}{sw}{coverage}"
            };

            if (sideWall)
            {
                var triangle = new Polyline(3)
                {
                    Layer = Layer,
                    Closed = true,
                    ColorIndex = ColorIndices.ByLayer,
                };

                double side = 12.0; // inches

                triangle.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
                triangle.AddVertexAt(1, new Point2d(side, (side/2)), 0, 0, 0);
                triangle.AddVertexAt(2, new Point2d(side, -(side / 2)), 0, 0, 0);

                record.AppendEntity(triangle);
            }
            else
            {
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
            }

            var attribute = new AttributeDefinition
            {
                Tag = "MODEL",
                TextString = "MODEL-123",
                Position = new Point3d(12, 4, 0),
                Height = 8.0,
                ColorIndex = ColorIndices.ByLayer,
                Layer = Layer,
            };

            if(sideWall)
            {
                attribute.Position = new Point3d(12, 10, 0);
            }    

            record.AppendEntity(attribute);

            var square = new Polyline(4)
            {
                Layer = CoverageLayer,
                Closed = true,
                ColorIndex = ColorIndices.ByLayer,
            };

            double coverageInches = coverage * 12;
            double radius = coverageInches / 2;

            if (sideWall)
            {
                square.AddVertexAt(0, new Point2d(0, radius), 0, 0, 0);
                square.AddVertexAt(1, new Point2d(coverageInches, radius), 0, 0, 0);
                square.AddVertexAt(2, new Point2d(coverageInches, -radius), 0, 0, 0);
                square.AddVertexAt(3, new Point2d(0, -radius), 0, 0, 0);
            }
            else
            {
                square.AddVertexAt(0, new Point2d(-radius, radius), 0, 0, 0);
                square.AddVertexAt(1, new Point2d(radius, radius), 0, 0, 0);
                square.AddVertexAt(2, new Point2d(radius, -radius), 0, 0, 0);
                square.AddVertexAt(3, new Point2d(-radius, -radius), 0, 0, 0);
            }

            record.AppendEntity(square);

            var text = new DBText
            { 
                 Layer = CoverageLayer,
                 Height = 16.0,
                 TextString = $"{coverage} X {coverage}",
                 Justify = AttachmentPoint.TopCenter,
                 AlignmentPoint = new Point3d(0, radius, 0),
                 ColorIndex = ColorIndices.ByLayer,
            };

            if(sideWall)
            {
                text.Justify = AttachmentPoint.TopLeft;
                text.AlignmentPoint = new Point3d(0, radius, 0);
            }

            record.AppendEntity(text);

            table.Add(record);

            return record;
        }
    }
}
