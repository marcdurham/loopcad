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

        [CommandMethod("LABELNODES")]
        public void LabelNodesCommand()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\nLabeling nodes...");

            Database db = HostApplicationServices.WorkingDatabase;
            Transaction trans = db.TransactionManager.StartTransaction();

            var headLabeler = new Labeler(trans, "HEADNUMBER", "HeadLabel2", "HeadLabels", ColorIndices.Magenta);
            var teeLabeler = new Labeler(trans, "TEENUMBER", "TeeLabel2", "TeeLabels", ColorIndices.Green);
            var pipeLabeler = new Labeler(trans, "PIPENUMBER", "PipeLabel2", "PipeLabels", ColorIndices.Blue)
            { 
                TextHeight = 3.0,
                XOffset = 0.0,
                YOffset = 0.0,
                HorizontalMode = TextHorizontalMode.TextCenter 
            };


            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
            BlockTableRecord modelSpace = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            int pipeNumber = 1;
            int nodeNumber = 1;
            int headNumber = 1;
            int teeNumber = 1;
            int riserNumber = 1;
            foreach (var objectId in modelSpace)
            {
                if (IsHead(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    headLabeler.CreateLabel($"H.{headNumber++}", block.Position);
                    //CreatePipeLabel(trans, pipeNumber, block.Position);
                }
                else if (IsTee(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    teeLabeler.CreateLabel($"T.{teeNumber++}", block.Position);
                }
                else if (IsPipe(trans, objectId))
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

        [CommandMethod("LABELPIPES")]
        public void LabelPipesCommand()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\nLabeling nodes...");

            Database db = HostApplicationServices.WorkingDatabase;
            Transaction trans = db.TransactionManager.StartTransaction();

            dynamic bt = db.BlockTableId;
            string str = "ABC";

            var textEnts =
                from btrs in (IEnumerable<dynamic>)bt
                from ent in (IEnumerable<dynamic>)btrs
                where
                ((ent.IsKindOf(typeof(DBText)) &&
                    (ent.TextString.Contains(str))) ||
                (ent.IsKindOf(typeof(MText)) &&
                    (ent.Contents.Contains(str))))
                select ent;

            Point3d pnt1 = new Point3d(0, 0, 0);
            Point3d pnt2 = new Point3d(10, 10, 0);

            Line lineObj = new Line(pnt1, pnt2);

            trans.AddNewlyCreatedDBObject(lineObj, true);
            trans.Commit();
        }
    }
}
