using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

namespace LoopCAD.WPF
{
    public class Initialization : IExtensionApplication
    {
        public void Initialize()
        {
           
        }

        public void Terminate()
        {
            
        }

        [CommandMethod("LABEL-NODES")]
        public void LabelNodesCommand()
        {
            Editor().WriteMessage("\nLabeling nodes...");
            
            int domesticTees = DomesticTeeLabeler.Run();
            Editor().WriteMessage($"\n{domesticTees} domestic tees labeled.");

            int heads = HeadLabeler.Run();
            Editor().WriteMessage($"\n{heads} tees labeled.");

            int tees = TeeLabeler.Run();
            Editor().WriteMessage($"\n{tees} heads labeled.");
        }

        [CommandMethod("LABEL-PIPES")]
        public void LabelPipesCommand()
        {
            Editor().WriteMessage("\nLabeling pipes...");

            int pipes = PipeLabeler.LabelAllPipes();

            Editor().WriteMessage($"\n{pipes} pipes labeled.");
        }

        [CommandMethod("RISER")]
        public void InsertRiserCommand()
        {
            Editor().WriteMessage("\nInserting riser...");

            var options = new PromptPointOptions(
                "Click location to insert riser")
            {
                AllowArbitraryInput = true,
            };

            Object osmode = Application.GetSystemVariable("OSMODE");
            Application.SetSystemVariable("OSMODE", 65);

            var result = Editor().GetPoint(options);

            if (result.Status == PromptStatus.OK)
            {
                RiserBuilder.Insert(result.Value);
            }

            Application.SetSystemVariable("OSMODE", osmode);
        }

        [CommandMethod("SAVE-DXF")]
        public void SaveAsDxfCommand()
        {
            Editor().WriteMessage("\nSaving as DXF...");

            string path = Application.DocumentManager.MdiActiveDocument.Name.Replace(".dwg", ".dxf");
            string directory = System.IO.Path.GetDirectoryName(path);
            string fileName = System.IO.Path.GetFileName(path);

            var dialog = new Microsoft.Win32.SaveFileDialog()
            {
                Filter = "Drawing EXchange Files (*.dxf)|*.dxf|All Files (*.*)|*.*",
                DefaultExt = ".dxf",
                FileName = fileName,
                InitialDirectory = directory
            };

            var result = dialog.ShowDialog();
            if (result.HasValue)
            {
                HostApplicationServices.WorkingDatabase.DxfOut(fileName: dialog.FileName, precision: 8, DwgVersion.Current);
                Editor().WriteMessage("\nDone");
            }
            else
            {
                Editor().WriteMessage("\nSave As DXF cancelled");
            }
        }

        [CommandMethod("GET-JOB-DATA")]
        public void GetJobDataCommand()
        {
            Editor().WriteMessage("\nGetting job data...");
            var data = JobData.Load();
            Editor().WriteMessage($"\nJob Number: {data.JobNumber}");
            Editor().WriteMessage("\nDone.");
        }


        [CommandMethod("IH2")]
        public void InsertHeadCommand()
        {
            Editor().WriteMessage("\nInserting head...");

            var options = new PromptPointOptions(
                "Click location to insert head")
            {
                AllowArbitraryInput = true,
            };

            object osmode = Application.GetSystemVariable("OSMODE");
            Application.SetSystemVariable("OSMODE", 65);

            using (var transaction = ModelSpace.StartTransaction())
            {
                var table = (BlockTable)transaction.GetObject(
                    Editor().Document.Database.BlockTableId, 
                    OpenMode.ForRead);

                var jigBlock = (BlockTableRecord)transaction.GetObject(
                    table["Head12"], 
                    OpenMode.ForRead);

                var jig = new BlockJig();
                Point3d point;
                PromptResult res = jig.DragMe(jigBlock.ObjectId, out point);

                if (res.Status == PromptStatus.OK)
                {
                    HeadBuilder.Insert(point);
                }

                transaction.Commit();
            }

            Application.SetSystemVariable("OSMODE", osmode);
        }

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
