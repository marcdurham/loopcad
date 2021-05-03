using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class JobData
    {
        public const string BlockName = "FloorConnector";
        public const string Layer = "Floor Connectors";
        readonly Database db;
        readonly DBDictionary table;
        readonly Transaction transaction;

        public JobData(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            table = (DBDictionary)transaction.GetObject(
                db.NamedObjectsDictionaryId,
                OpenMode.ForRead);
        }

        public static void Test()
        {
            using (var trans = ModelSpace.StartTransaction())
            {
                new JobData(trans).StartTest();
            }
        }

        void StartTest()
        {
            if(!table.Contains("job_data"))
            {
                Editor().WriteMessage("\nJob data does not exist in this file");
            }

            var dict = transaction.GetObject(table.GetAt("job_data"), OpenMode.ForRead) as DBDictionary;

            if(!dict.Contains("job_number"))
            {
                Editor().WriteMessage("\nJob number does not exist");
            }

            var xrec = (Xrecord)transaction.GetObject(dict.GetAt("job_number"), OpenMode.ForRead);

            string jobNumber = string.Empty;
            foreach (TypedValue value in xrec.Data.AsArray())
            {
                if (value.TypeCode == 1)
                {
                    jobNumber = value.Value as string;
                }
            }

            Editor().WriteMessage($"Job number is {jobNumber}");
        }

        Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
