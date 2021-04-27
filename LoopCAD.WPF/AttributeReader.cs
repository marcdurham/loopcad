using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class AttributeReader
    {
        public static string TextString(Transaction transaction, ObjectId blockRefObjectId, string blockName, string tag)
        {
            var blockRef = transaction.GetObject(blockRefObjectId, OpenMode.ForRead) as BlockReference;
            if (blockRef == null)
            {
                return null;
            }

            foreach(ObjectId attId in blockRef.AttributeCollection)
            {
                var attRef = transaction.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if(string.Equals(attRef.Tag, tag, System.StringComparison.OrdinalIgnoreCase))
                {
                    return attRef.TextString;
                }    
            }

            return null;
        }
    }
}
