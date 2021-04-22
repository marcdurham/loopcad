using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class LayerCreator
    {
        public static void Ensure(string name, short colorIndex)
        {
            using(var transaction = ModelSpace.StartTransaction())
            using (var table = (LayerTable)transaction
                .GetObject(
                    HostApplicationServices.WorkingDatabase.LayerTableId,
                    OpenMode.ForRead))
            {

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
                transaction.Commit();
            }
        }
    }
}
