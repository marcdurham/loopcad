using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class LayerCreator
    {
        readonly Transaction transaction;

        public LayerCreator(Transaction transaction)
        {
            this.transaction = transaction;
        }

        public static void Ensure(
            Transaction transaction, 
            string name, 
            short colorIndex)
        {
            new LayerCreator(transaction).Ensure(name, colorIndex);
        }

        public void Ensure(string name, short colorIndex = ColorIndices.White)
        {
            var table = (LayerTable)transaction
                .GetObject(
                    HostApplicationServices.WorkingDatabase.LayerTableId, 
                    OpenMode.ForRead);

            if (table.Has(name))
            {
                return;
            }
            
            var record = new LayerTableRecord()
            {
                Name = name,
                Color = Color.FromColorIndex(ColorMethod.ByLayer, colorIndex)
            };

            table.UpgradeOpen();
            table.Add(record);

            transaction.AddNewlyCreatedDBObject(record, true);
        }
    }
}
