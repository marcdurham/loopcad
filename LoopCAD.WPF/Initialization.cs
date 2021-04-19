﻿using Autodesk.AutoCAD.ApplicationServices;
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

                    CreateNodeLabel(trans, $"H.{headNumber++}", block.Position);
                    //CreatePipeLabel(trans, pipeNumber, block.Position);
                }
                else if (IsTee(trans, objectId))
                {
                    var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;

                    CreateNodeLabel(trans, $"T.{teeNumber++}", block.Position);
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

        static void CreatePipeLabel(Transaction trans, int pipeNumber, Point3d position)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
            BlockTableRecord modelSpace = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            var pipeLabel = new DBText()
            {
                Layer = "Heads",
                ColorIndex = 150,
                TextString = $"p{pipeNumber}",
                Position = position
            };

            modelSpace.AppendEntity(pipeLabel);
            trans.AddNewlyCreatedDBObject(pipeLabel, true);
        }

        static void CreateNodeLabel(Transaction trans, string text, Point3d position)
        {
            Database db = HostApplicationServices.WorkingDatabase;
      
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

            BlockTableRecord modelSpace = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


            BlockTableRecord nodeLabelDef = NodeLabel(trans);
            var newBlockRef = new BlockReference(position, nodeLabelDef.Id);

            modelSpace.AppendEntity(newBlockRef);
            trans.AddNewlyCreatedDBObject(newBlockRef, true);

            // Iterate block definition to find all non-constant
            // AttributeDefinitions
            foreach (ObjectId id in nodeLabelDef)
            {
                DBObject obj = id.GetObject(OpenMode.ForRead);
                var attDef = obj as AttributeDefinition;

                if ((attDef != null) && (!attDef.Constant) && attDef.Tag.ToUpper() == "NODENUMBER")
                {
                    //This is a non-constant AttributeDefinition
                    //Create a new AttributeReference
                    using (var attRef = new AttributeReference())
                    {
                        attRef.SetAttributeFromBlock(attDef, newBlockRef.BlockTransform);
                        attRef.TextString = text;
                        //Add the AttributeReference to the BlockReference
                        newBlockRef.AttributeCollection.AppendAttribute(attRef);
                        trans.AddNewlyCreatedDBObject(attRef, true);
                    }
                }
            }
        }

        static BlockTableRecord NodeLabel(Transaction trans)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
            BlockTableRecord nodeLabelDef;
            if (!blkTbl.Has("NodeLabel"))
            {
                nodeLabelDef = NodeLabelDefFrom(trans, blkTbl);
                trans.AddNewlyCreatedDBObject(nodeLabelDef, true);
            }
            else
            {
                //nodeLabelId = blkTbl["NodeLabel"];
                nodeLabelDef = trans.GetObject(blkTbl["NodeLabel"], OpenMode.ForRead) as BlockTableRecord;
                //if(nodeLabelDef.Id == nodeLabelId)
                //{
                //    Debug.WriteLine("Yep");
                //}
            }

            return nodeLabelDef;
        }

        static BlockTableRecord NodeLabelDefFrom(Transaction trans, BlockTable blkTbl)
        {
            var nodeLabelDef = new BlockTableRecord();
            nodeLabelDef.Name = "NodeLabel";

            var attRef = new AttributeDefinition();
            attRef.Height = 4.75;
            attRef.TextStyleId = ArialStyle(trans);


            attRef.Tag = "NODENUMBER";
            attRef.TextString = "N.99";
            attRef.Position = new Point3d(6, -6, 0);
            nodeLabelDef.AppendEntity(attRef);
            blkTbl.Add(nodeLabelDef);
            
            return nodeLabelDef;
        }

        static ObjectId ArialStyle(Transaction trans)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            var styles = trans.GetObject(db.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;
            var currentStyle = trans.GetObject(db.Textstyle, OpenMode.ForWrite) as TextStyleTableRecord;

            var style = new TextStyleTableRecord()
            {
                Font = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(
                    "ARIAL",
                    bold: false,
                    italic: false,
                    characters: currentStyle.Font.CharacterSet,
                    pitchAndFamily: currentStyle.Font.PitchAndFamily),
            };

            ObjectId styleId = styles.Add(style);
            
            trans.AddNewlyCreatedDBObject(style, true);
            
            return styleId;
        }

        [CommandMethod("LABELSTUFF")]
        public void LabelStuffCommand()
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

            //msBlkRec.AppendEntity(lineObj);
            trans.AddNewlyCreatedDBObject(lineObj, true);
            trans.Commit();

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

            //msBlkRec.AppendEntity(lineObj);
            trans.AddNewlyCreatedDBObject(lineObj, true);
            trans.Commit();

        }
    }
}