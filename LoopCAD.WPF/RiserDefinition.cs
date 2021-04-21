using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoopCAD.WPF
{
    public class RiserDefinition
    {
        const string blockName = "Riser";
        readonly BlockTableRecord definition;
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

            definition = DefinitionFrom(null);
        }

        public static BlockTableRecord ExistingOrNew(Transaction transaction)
        {
            return new RiserDefinition(transaction)
                .ExistingOrNewDefinition();
        }
        
        BlockTableRecord ExistingOrNewDefinition()
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
            var record = new BlockTableRecord
            {
                Name = blockName
            };

            var circle = new Circle()
            {
                Center = new Point3d(0, 0, 0),
                Radius = 9.0, // inches
                Layer = "Risers",
                ColorIndex = ColorIndices.ByLayer,
            };

            record.AppendEntity(circle);
            table.Add(record);

            return record;
        }
    }
}
