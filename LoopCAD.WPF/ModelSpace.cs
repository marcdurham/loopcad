using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class ModelSpace
    {
        public static BlockTableRecord From(Transaction transaction)
        {
            BlockTable table = transaction.GetObject(
                HostApplicationServices.WorkingDatabase.BlockTableId,
                OpenMode.ForWrite) as BlockTable;

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
    }
}
