using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace LoopCAD.WPF
{
    public class Initialization : IExtensionApplication
    {
        public void Initialize()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            Editor().WriteMessage($"\nLoopCAD .NET module version {version} loaded.");
        }

        public void Terminate()
        {
            
        }

        [CommandMethod("ELEVATION-BOX")]
        public void  ElevationBoxCommand()
        {
            Editor().WriteMessage("\nDrawing elevation box...");

            var firstCornerOptions = new PromptPointOptions(
              "Click first corner of elevation box")
            {
                AllowArbitraryInput = true,
            };

            PromptPointResult first = Editor().GetPoint(firstCornerOptions);

            PromptPointResult other = Editor()
                .GetCorner("Click other corner of the elevation box", first.Value);

            PromptDoubleOptions elevation = new PromptDoubleOptions("Enter elevation in feet")
            {
                DefaultValue = 100.0
            };

            PromptDoubleResult elevationResult = Editor()
                .GetDouble(elevation);

            ElevationBoxBuilder.Start(
                first.Value,
                other.Value,
                elevationResult.Value);
           
            Editor().WriteMessage("\nDone drawing elevation box.");
        }


        [CommandMethod("FLOOR-TAG")]
        public void InsertFloorTagCommand()
        {
            Editor().WriteMessage("\nInserting floor tag...");

            var options = new PromptPointOptions(
                "Click location to insert floor tag")
            {
                AllowArbitraryInput = true,
            };

            Object osmode = Application.GetSystemVariable("OSMODE");
            Application.SetSystemVariable("OSMODE", 65);

            var result = Editor().GetPoint(options);

            if (result.Status == PromptStatus.OK)
            {
                FloorTagBuilder.Insert(result.Value);
            }

            Application.SetSystemVariable("OSMODE", osmode);
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
            data.CalculatedByCompanies = new List<string>() { 
                "X-Fire", 
                "13dpex.com", 
                "Testing Company", 
                "Other Company" };
            data.CalculatedByCompany = "Other Company";
            
            data.SprinklerPipeTypes = new List<string> {
                "Rehau Brass",
                "Rehau Plastic",
                "Other"};
            data.SprinklerPipeType = "Other";

            data.SprinklerFittingTypes = new List<string>() {
                "Rehau PEX",
                "Copper",
                "CPVC",
                "Other"};
            data.SprinklerFittingType = "CPVC";

            data.SupplyPipeTypes = new List<string>() {
                "Poly",
                "Rehau PEX",
                "Spears Flameguard CPVPC",
                "Copper",
                "CPVC",
                "Other"
            };
            data.SupplyPipeType = "Spears Flameguard CPVPC";

            Editor().WriteMessage($"\nJob Number: {data?.JobNumber}");
            Editor().WriteMessage($"\nJob Name: {data?.JobName}");
            Editor().WriteMessage($"\nSite Location: {data?.JobSiteAddress}");
            Editor().WriteMessage($"\nSupply Static Pressure: {data?.SupplyStaticPressure}");
            Editor().WriteMessage($"\nDone.");

            var form = new JobDataForm(data);
            form.JobData = data;
            form.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            form.ShowDialog();

            Editor().WriteMessage("\nChanges:");
            Editor().WriteMessage($"\nJob Number: {data?.JobNumber}");
            Editor().WriteMessage($"\nJob Name: {data?.JobName}");
            Editor().WriteMessage($"\nSite Location: {data?.JobSiteAddress}");
            Editor().WriteMessage($"\nSupply Static Pressure: {data?.SupplyStaticPressure}");
            Editor().WriteMessage("\nDone.");
        }

        [CommandMethod("H20")]
        public void InsertHead20Command()
        {
            InsertHead(20);
        }

        [CommandMethod("H18")]
        public void InsertHead18Command()
        {
            InsertHead(18);
        }

        [CommandMethod("H16")]
        public void InsertHead16Command()
        {
            InsertHead(16);
        }

        [CommandMethod("H14")]
        public void InsertHead14Command()
        {
            InsertHead(14);
        }
        
        [CommandMethod("H12")]
        public void InsertHead12Command()
        {
            InsertHead(12);
        }

        [CommandMethod("SW12")]
        public void InsertSideWallHead12Command()
        {
            InsertHeadSideWall(12);
        }

        [CommandMethod("SW14")]
        public void InsertSideWallHead14Command()
        {
            InsertHeadSideWall(14);
        }

        [CommandMethod("SW16")]
        public void InsertSideWallHead16Command()
        {
            InsertHeadSideWall(16);
        }

        [CommandMethod("SW18")]
        public void InsertSideWallHead18Command()
        {
            InsertHeadSideWall(18);
        }

        [CommandMethod("SW20")]
        public void InsertSideWallHead20Command()
        {
            InsertHeadSideWall(20);
        }

        static void InsertHead(int coverage)
        {
            Editor().WriteMessage("\nInserting head...");

            ObjectId layerId = Layer.Ensure("HeadCoverage", ColorIndices.Yellow);
            Layer.Show(layerId);
            object osmode = Application.GetSystemVariable("OSMODE");
            Application.SetSystemVariable("OSMODE", 65);

            HeadBuilder.Insert(coverage);

            Layer.Hide(layerId);
            Application.SetSystemVariable("OSMODE", osmode);
        }

        static void InsertHeadSideWall(int coverage)
        {
            Editor().WriteMessage("\nInserting head...");

            ObjectId layerId = Layer.Ensure("HeadCoverage", ColorIndices.Yellow);
            Layer.Show(layerId);
            object osmode = Application.GetSystemVariable("OSMODE");
            object orthomode = Application.GetSystemVariable("ORTHOMODE");
            Application.SetSystemVariable("OSMODE", 65);
            Application.SetSystemVariable("ORTHOMODE", 1);

            var pointResult = Editor().GetPoint("Where to place head");
            if(pointResult.Status != PromptStatus.OK)
            {
                return;
            }

            var options = new PromptAngleOptions("Angle of sidewall spray (press F8 to turn ORTHO off)")
            {
                BasePoint = pointResult.Value,
                UseBasePoint = true
            };

            var angleResult = Editor().GetAngle(options);

            if (angleResult.Status == PromptStatus.OK)
            {
                HeadBuilder.Insert(coverage, pointResult.Value, angleResult.Value);
            }

            Layer.Hide(layerId);
            Application.SetSystemVariable("OSMODE", osmode);
            Application.SetSystemVariable("ORTHOMODE", orthomode);
        }

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    }
}
