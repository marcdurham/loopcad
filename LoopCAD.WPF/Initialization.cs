using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        [CommandMethod("TESTLABELNODES")]
        public void TestLabelNodesCommand()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\nI have created my first command");
            
            Database db = HostApplicationServices.WorkingDatabase;
            Transaction trans = db.TransactionManager.StartTransaction();

            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord msBlkRec = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            Point3d pnt1 = new Point3d(0, 0, 0);
            //Point3d pnt2 = new Point3d(10, 10, 0);
            //Point3d pnt2 = ed.GetPoint("Give me another point").Value;
            PromptPointOptions prPtOpt = new PromptPointOptions("\nSpecify start point: ");
            prPtOpt.AllowArbitraryInput = false;
            prPtOpt.AllowNone = true;

            PromptPointResult prPtRes1 = ed.GetPoint(prPtOpt);
            if (prPtRes1.Status != PromptStatus.OK) return;
            Point3d pnt2 = prPtRes1.Value;

            Line lineObj = new Line(pnt1, pnt2);
            msBlkRec.AppendEntity(lineObj);
            trans.AddNewlyCreatedDBObject(lineObj, true);
            trans.Commit();

        }

        [CommandMethod("LABEL-NODES")]
        public void LabelNodesCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("\nLabeling nodes...");

            Transaction trans = StartTransaction();

            var headLabeler = new Labeler(trans, "HEADNUMBER", "HeadLabel2", "HeadLabels", ColorIndices.Magenta);
            var teeLabeler = new Labeler(trans, "TEENUMBER", "TeeLabel2", "TeeLabels", ColorIndices.Green);

            int headNumber = 1;
            int teeNumber = 1;
            int riserNumber = 1;
            foreach (var objectId in ModelSpace(trans))
            {
                if (IsHead(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    headLabeler.CreateLabel($"H.{headNumber++}", block.Position);
                }
                else if (IsTee(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    teeLabeler.CreateLabel($"T.{teeNumber++}", block.Position);
                }
            }

            editor.WriteMessage($"\n{headNumber} heads labeled.");
            editor.WriteMessage($"\n{teeNumber} tees labeled.");
            trans.Commit();
        }

        [CommandMethod("LABEL-PIPES")]
        public void LabelPipesCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("\nLabeling pipes...");

            Transaction trans = StartTransaction();

            var pipeLabeler = new Labeler(trans, "PIPENUMBER", "PipeLabel2", "PipeLabels", ColorIndices.Blue)
            {
                TextHeight = 3.0,
                XOffset = 0.0,
                YOffset = 0.0,
                HorizontalMode = TextHorizontalMode.TextCenter
            };

            int pipeNumber = 1;
            foreach (var objectId in ModelSpace(trans))
            {
                if (IsPipe(trans, objectId))
                {
                    var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;

                    for (int i = 1; i < polyline.NumberOfVertices; i++)
                    {
                        Point3d lastVertex = polyline.GetPoint3dAt(i - 1);
                        Point3d vertex = polyline.GetPoint3dAt(i);
                        pipeLabeler.CreateLabel(
                            text: $"p{pipeNumber}",
                            position: Midpoint(vertex, lastVertex));
                    }

                    pipeNumber++;
                }
            }

            trans.Commit();
            editor.WriteMessage($"\n{pipeNumber} pipes labeled.");
        }

        static Transaction StartTransaction()
        {
            return HostApplicationServices
                .WorkingDatabase
                .TransactionManager
                .StartTransaction();
        }

        static BlockTableRecord ModelSpace(Transaction trans)
        {
            BlockTable blkTbl = trans.GetObject(
                HostApplicationServices.WorkingDatabase.BlockTableId, 
                OpenMode.ForWrite) as BlockTable;
            
            BlockTableRecord modelSpace = trans.GetObject(
                blkTbl[BlockTableRecord.ModelSpace], 
                OpenMode.ForWrite) as BlockTableRecord;

            return modelSpace;
        }


        bool IsHead(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" && 
                string.Equals(block.Layer, "Heads", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("HEAD");
        }

        bool IsTee(Transaction trans, ObjectId objectId)
        {
            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, "Tees", StringComparison.OrdinalIgnoreCase) &&
                block.Name.ToUpper().StartsWith("TEE");
        }

        bool IsPipe(Transaction trans, ObjectId objectId)
        {
            var polyline = trans.GetObject(objectId, OpenMode.ForRead) as Polyline;
            return objectId.ObjectClass.DxfName == "LWPOLYLINE" &&
                string.Equals(polyline.Layer, "Pipes", StringComparison.OrdinalIgnoreCase);
        }

        Point3d Midpoint(Point3d a, Point3d b)
        {
            return new Point3d(
                x: (b.X - a.X) / 2 + a.X,
                y: (b.Y - a.Y) / 2 + a.Y,
                z: 0);
        }
    }
}
