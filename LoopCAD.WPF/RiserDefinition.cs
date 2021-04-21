using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class RiserDefinition
    {
        const string blockName = "Riser";
        const string layer = "Risers";
        readonly Database db;
        readonly BlockTable table;

        Transaction transaction;

        public RiserDefinition(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            table = transaction.GetObject(
                db.BlockTableId,
                OpenMode.ForWrite) as BlockTable;
        }

        public static BlockTableRecord Define(Transaction transaction)
        {
            return new RiserDefinition(transaction)
                .Ensure();
        }
        
        BlockTableRecord Ensure()
        {
            BlockTableRecord record;

            if (!table.Has(blockName))
            {
                record = DefinitionFrom(table);
                transaction.AddNewlyCreatedDBObject(record, true);
            }
            else
            {
                record = transaction.GetObject(table[blockName], OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }

        BlockTableRecord DefinitionFrom(BlockTable table)
        {
            LayerCreator.Ensure(
                transaction, 
                name: layer, 
                colorIndex: ColorIndices.Cyan);

            var record = new BlockTableRecord
            {
                Name = blockName
            };

            var circle = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 9.0, // inches
                Layer = layer,
                ColorIndex = ColorIndices.ByLayer,
            };

            record.AppendEntity(circle);
            table.Add(record);

            return record;
        }
    }
}
