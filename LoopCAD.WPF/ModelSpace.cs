using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class ModelSpace
    {
        public static BlockTableRecord From(Transaction transaction)
        {
            BlockTable table = transaction.GetObject(
                Database().BlockTableId,
                OpenMode.ForRead) as BlockTable;

            BlockTableRecord modelSpace = transaction.GetObject(
                table[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

            return modelSpace;
        }

        public static Transaction StartTransaction()
        {
            return HostApplicationServices
                .WorkingDatabase
                .TransactionManager
                .StartTransaction();
        }

        public static Database Database()
        {
            return HostApplicationServices.WorkingDatabase;
        }
    }
}
