using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class LabelSpecs
    {
        public string Tag { get; set; } = "NUMBER";
        public string BlockName { get; set; } = "Label";
        public string Layer { get; set; } = "Labels";
        public short LayerColorIndex { get; set; } = ColorIndices.ByLayer;
        public TextHorizontalMode HorizontalMode { get; set; }
        public double TextHeight { get; set; } = 8.0;
        public double XOffset { get; set; } = 10.0;
        public double YOffset { get; set; } = -10.0;
    }

    public class Labeler
    {
        static ObjectId labelBlockDefId = ObjectId.Null;

        readonly LabelSpecs specs = new LabelSpecs();

        public Labeler(LabelSpecs labelSpecs) //string tag, string blockName, string layer, short layerColorIndex)
        {
            specs = labelSpecs;

            Layer.Ensure(specs.Layer, specs.LayerColorIndex);

            //arialStyle = ArialStyle();
            labelBlockDefId = ExistingOrNewLabelDefId();
        }

        public void CreateLabel(string text, Point3d position)
        {
            NewNodeLabel(text, position);
        }

        void NewNodeLabel(string text, Point3d position)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var table = transaction.GetObject(
                HostApplicationServices.WorkingDatabase.BlockTableId, 
                OpenMode.ForWrite) as BlockTable)
            using (var modelSpace = transaction.GetObject(
                table[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord)
            {
                using(var labelBlockDef = transaction.GetObject(
                    labelBlockDefId, 
                    OpenMode.ForRead) as BlockTableRecord)
                using (var blockRef = new BlockReference(position, labelBlockDefId)
                {
                    Layer = specs.Layer,
                    ColorIndex = ColorIndices.ByLayer
                })
                {
                    modelSpace.AppendEntity(blockRef);
                    transaction.AddNewlyCreatedDBObject(blockRef, true);

                    foreach (ObjectId id in labelBlockDef)
                    {
                        using (var def = id.GetObject(OpenMode.ForRead) as AttributeDefinition)
                        {
                            if ((def != null) && (!def.Constant) 
                                && def.Tag.ToUpper() == specs.Tag)
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

                    transaction.Commit();
                }
            }
        }

        ObjectId ExistingOrNewLabelDefId()
        {
            using (var trans = ModelSpace.StartTransaction())
            using (var tab = (BlockTable)trans.GetObject(
                HostApplicationServices.WorkingDatabase.BlockTableId,
                OpenMode.ForRead))
            {
                if (tab.Has(specs.BlockName))
                {
                    using (var existing = tab[specs.BlockName].GetObject(OpenMode.ForRead) as BlockTableRecord)
                    {
                        if (!PropertiesMatch(trans, existing))
                        {
                            existing.UpgradeOpen();
                            existing.Erase(true);
                            trans.Commit();
                        }
                        else
                        {
                            return existing.Id;
                        }
                    }
                }
            }

            using (var trans = ModelSpace.StartTransaction())
            using (var tab = (BlockTable)trans.GetObject(
                HostApplicationServices.WorkingDatabase.BlockTableId,
                OpenMode.ForRead))
            {

                BlockTableRecord record = NewLabelDefIdFrom();

                tab.UpgradeOpen();
                tab.Add(record);
                trans.AddNewlyCreatedDBObject(record, true);

                var attDef = NewAttributeDef();
                record.AppendEntity(attDef);
                trans.AddNewlyCreatedDBObject(attDef, true);

                trans.Commit();
                return record.Id;
            }
        }

        bool PropertiesMatch(Transaction transaction, BlockTableRecord record)
        {
            var attDef = AttributeReader.AttributeDefWithTagNamed(
                transaction, 
                record.Id, 
                specs.Tag);

            return attDef != null &&
                attDef.Position.X == specs.XOffset &&
                attDef.Position.Y == specs.YOffset &&
                attDef.Height == specs.TextHeight &&
                attDef.Layer == specs.Layer &&
                attDef.ColorIndex == ColorIndices.ByLayer;
        }

        BlockTableRecord NewLabelDefIdFrom()
        {
            var record = new BlockTableRecord
            {
                Name = specs.BlockName
            };

            return record;
        }

        AttributeDefinition NewAttributeDef()
        {
            var definition = new AttributeDefinition()
            {
                Height = specs.TextHeight,
                //TextStyleId = arialStyle,
                Layer = specs.Layer,
                ColorIndex = ColorIndices.ByLayer,
                Tag = specs.Tag,
                TextString = "X.99",
                Position = new Point3d(specs.XOffset, specs.YOffset, 0),
                HorizontalMode = specs.HorizontalMode,
            };

            return definition;
        }

        static ObjectId ArialStyle()
        {
            using(var transaction = ModelSpace.StartTransaction())
            using (var styles = (TextStyleTable)transaction.GetObject(
                HostApplicationServices.WorkingDatabase.TextStyleTableId,
                OpenMode.ForWrite))
            using (var currentStyle = (TextStyleTableRecord)transaction.GetObject(
                HostApplicationServices.WorkingDatabase.Textstyle,
                OpenMode.ForWrite))
            {
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
                transaction.Commit();

                return styleId;
            }
        }
    }
}
