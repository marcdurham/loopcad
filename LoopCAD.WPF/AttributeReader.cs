using Autodesk.AutoCAD.DatabaseServices;

namespace LoopCAD.WPF
{
    public class AttributeReader
    {
        public static string TextString(
            Transaction transaction,
            ObjectId blockRefObjectId, 
            string tag)
        {
            var attRef = AttributeWithTagNamed(transaction, blockRefObjectId, tag);

            return attRef?.TextString;
        }

        public static bool HasAttributeWithTag(
            Transaction transaction,
            ObjectId blockRefObjectId,
            string tag)
        {
            return AttributeWithTagNamed(transaction, blockRefObjectId, tag)
                != null;
        }

        public static bool HasAttributeDefWithTag(
            Transaction transaction,
            ObjectId blockRefObjectId,
            string tag)
        {
            return AttributeDefWithTagNamed(transaction, blockRefObjectId, tag)
                != null;
        }

        public static AttributeDefinition AttributeDefWithTagNamed(
            Transaction transaction,
            ObjectId blockDefObjectId,
            string tag)
        {
            var blockDef = transaction.GetObject(blockDefObjectId, OpenMode.ForRead) as BlockTableRecord;
            if (blockDef == null)
            {
                return null;
            }

            foreach (ObjectId attId in blockDef)
            {
                var attDef = transaction.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                if (attDef != null && string.Equals(attDef.Tag, tag, System.StringComparison.OrdinalIgnoreCase))
                {
                    return attDef;
                }
            }

            return null;
        }

        static AttributeReference AttributeWithTagNamed(
           Transaction transaction,
           ObjectId blockRefObjectId,
           string tag)
        {
            var blockRef = transaction.GetObject(blockRefObjectId, OpenMode.ForRead) as BlockReference;
            if (blockRef == null)
            {
                return null;
            }

            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                var attRef = transaction.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (string.Equals(attRef.Tag, tag, System.StringComparison.OrdinalIgnoreCase))
                {
                    return attRef;
                }
            }

            return null;
        }
    }
}
