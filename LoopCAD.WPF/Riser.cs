﻿using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class Riser
    {
        public const string BlockName = "FloorConnector";
        public const string Layer = "Floor Connectors";
        readonly Database db;
        readonly BlockTable table;
        readonly Transaction transaction;

        public Riser(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            table = transaction.GetObject(
                db.BlockTableId,
                OpenMode.ForWrite) as BlockTable;
        }

        public static void Insert(Point3d position)
        {
            new Riser(StartTransaction())
                .InsertAt(position);
        }

        void InsertAt(Point3d position)
        {
            var labelBlockDef = new BlockReference(position, Define().Id)
            {
                Layer = "Floor Connectors",
                ColorIndex = ColorIndices.ByLayer
            };

            ModelSpace.From(transaction).AppendEntity(labelBlockDef);
            transaction.AddNewlyCreatedDBObject(labelBlockDef, true);

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
            WPF.Layer.Ensure(Layer, ColorIndices.Cyan);

            var record = new BlockTableRecord
            {
                Name = BlockName
            };

            var circle = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 9.0, // inches
                Layer = Layer,
                ColorIndex = ColorIndices.ByLayer,
            };

            record.AppendEntity(circle);
            table.Add(record);

            return record;
        }
    }
}
