﻿using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class Labeler
    {
        readonly Transaction transaction;
        readonly Database db;
        BlockTable table;
        BlockTableRecord modelSpace;
        string tag = "";
        string blockName = "";
        string layer = "";

        public Labeler(Transaction transaction, string tag, string blockName, string layer, short layerColorIndex)
        {
            this.transaction = transaction;
            db = HostApplicationServices.WorkingDatabase;

            this.tag = tag;
            this.blockName = blockName;
            this.layer = layer;

            LayerCreator.Ensure(layer, layerColorIndex);
        }

        public double TextHeight { get; set; } = 4.75;
        public double XOffset { get; set; } = 6;
        public double YOffset { get; set; } = -6;
        public TextHorizontalMode HorizontalMode { get; set; }

        public void CreateLabel(string text, Point3d position)
        {
            NewNodeLabel(text, position);
        }

        void NewNodeLabel(string text, Point3d position)
        {
            using (table = transaction.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable)
            {
                modelSpace = transaction.GetObject(
                    table[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                BlockTableRecord record = ExistingOrNewLabelDef();

                var blockRef = new BlockReference(position, record.Id)
                {
                    Layer = layer,
                    ColorIndex = ColorIndices.ByLayer
                };

                modelSpace.AppendEntity(blockRef);
                transaction.AddNewlyCreatedDBObject(blockRef, true);

                foreach (ObjectId id in record)
                {
                    var def = id.GetObject(OpenMode.ForRead) as AttributeDefinition;

                    if ((def != null) && (!def.Constant) && def.Tag.ToUpper() == tag)
                    {
                        using (var ar = new AttributeReference())
                        {
                            ar.SetAttributeFromBlock(def, blockRef.BlockTransform);
                            ar.TextString = text;

                            blockRef.AttributeCollection.AppendAttribute(ar);
                            transaction.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                }
            }
        }

        BlockTableRecord ExistingOrNewLabelDef()
        {
            BlockTableRecord record;

            if (!table.Has(blockName))
            {
                record = LabelDefFrom(table);
                transaction.AddNewlyCreatedDBObject(record, true);
            }
            else
            {
                record = transaction.GetObject(table[blockName], OpenMode.ForRead) as BlockTableRecord;
            }

            return record;
        }

        BlockTableRecord LabelDefFrom(BlockTable table)
        {
            var record = new BlockTableRecord
            {
                Name = blockName
            };

            var definition = new AttributeDefinition()
            {
                Height = TextHeight,
                TextStyleId = ArialStyle(),
                Layer = layer,
                ColorIndex = ColorIndices.ByLayer,
                Tag = tag,
                TextString = "X.99",
                Position = new Point3d(XOffset, YOffset, 0),
                HorizontalMode = HorizontalMode,
            };

            record.AppendEntity(definition);
            table.Add(record);

            return record;
        }

        ObjectId ArialStyle()
        {
            var styles = (TextStyleTable)transaction.GetObject(
                db.TextStyleTableId, 
                OpenMode.ForWrite);

            var currentStyle = (TextStyleTableRecord)transaction.GetObject(
                db.Textstyle, 
                OpenMode.ForWrite);

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

            transaction.AddNewlyCreatedDBObject(style, true);

            return styleId;
        }
    }
}
