using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

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
            
            int heads = HeadLabeler.Run();
            Editor().WriteMessage($"\n{heads} tees labeled.");

            int tees = TeeLabeler.Run();
            Editor().WriteMessage($"\n{tees} heads labeled.");
        }

        [CommandMethod("LABEL-PIPES")]
        public void LabelPipesCommand()
        {
            Editor().WriteMessage("\nLabeling pipes...");

            int pipes = PipeLabeler.Run();

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

            var result = Editor().GetPoint(options);

            if (result.Status == PromptStatus.OK)
            {
                RiserBuilder.Insert(result.Value);
            }
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

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
