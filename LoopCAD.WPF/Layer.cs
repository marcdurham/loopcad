using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class Layer
    {
        public static ObjectId Ensure(string name, short colorIndex)
        {
            using(var transaction = ModelSpace.StartTransaction())
            using(var table = (LayerTable)transaction
                .GetObject(
                    HostApplicationServices.WorkingDatabase.LayerTableId,
                    OpenMode.ForRead))
            {
                if (table.Has(name))
                {
                    var layer = transaction.GetObject(
                        table[name],
                        OpenMode.ForWrite) as LayerTableRecord;

                    layer.Color = Color.FromColorIndex(ColorMethod.ByLayer, colorIndex);
                    transaction.Commit();
                    return table[name];
                }

                var record = new LayerTableRecord()
                {
                    Name = name,
                    Color = Color.FromColorIndex(ColorMethod.ByLayer, colorIndex),
                };

                table.UpgradeOpen();
                ObjectId id = table.Add(record);

                transaction.AddNewlyCreatedDBObject(record, true);
                transaction.Commit();

                return id;
            }
        }

        public static void Show(string name)
        {
            HideShow(name, false);
        }

        public static void Hide(string name)
        {
            HideShow(name, true);
        }
        
        public static void Show(ObjectId id)
        {
            HideShow(id, false);
        }

        public static void Hide(ObjectId id)
        {
            HideShow(id, true);
        }

        public static void HideShow(string name, bool isHidden)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var table = (LayerTable)transaction
                .GetObject(
                    HostApplicationServices.WorkingDatabase.LayerTableId,
                    OpenMode.ForRead))
            {

                if (table.Has(name))
                {
                    var layer = transaction.GetObject(
                        table[name],
                        OpenMode.ForWrite) as LayerTableRecord;

                    layer.IsHidden = isHidden;
                    transaction.Commit();
                }
            }
        }

        static void HideShow(ObjectId id, bool isHidden)
        {
            using(var transaction = ModelSpace.StartTransaction())
            {
                var layer = transaction.GetObject(
                    id,
                    OpenMode.ForWrite) as LayerTableRecord;
                
                layer.IsOff = isHidden;

                transaction.Commit();
            }
        }
    }
}
