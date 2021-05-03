using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class JobData
    {
        readonly Database db;
        readonly DBDictionary namedObjectDict;
        readonly Transaction transaction;

        public JobData(Transaction transaction)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;
            namedObjectDict = (DBDictionary)transaction.GetObject(
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
            string jobNumber = NamedObjectDictionary
                .KeyValue("job_data", "job_number");

            Editor().WriteMessage($"Job number is {jobNumber}");
        }

        Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
