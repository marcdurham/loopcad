using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class NamedObjectDictionary
    {
        public static bool KeyValue(string dictionaryName, string key, out string value)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var db = HostApplicationServices.WorkingDatabase)
            using (var namedObjectDict = (DBDictionary)transaction.GetObject(
                    db.NamedObjectsDictionaryId,
                    OpenMode.ForRead))
            {
                if (!namedObjectDict.Contains(dictionaryName))
                {
                    throw new Exception(
                        $"The dictionary named '{dictionaryName}' does not exist in this DWG file");
                }

                var dictionary = transaction.GetObject(
                    namedObjectDict.GetAt(dictionaryName),
                    OpenMode.ForRead) as DBDictionary;

                if (!dictionary.Contains(key))
                {
                    value = null;
                    return false;
                    throw new Exception(
                        $"Key '{key}' does not exist in the dictionary '{dictionaryName}'");
                }

                var xRecord = (Xrecord)transaction.GetObject(
                    dictionary.GetAt(key),
                    OpenMode.ForRead);

                value = string.Empty;
                foreach (TypedValue typedValue in xRecord.Data.AsArray())
                {
                    if (typedValue.TypeCode == 1)
                    {
                        value = typedValue.Value as string;
                    }
                }

                return true;
            }
        }

        /*
        public static void SetKeyValue(string dictionaryName, string key, string value)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var db = HostApplicationServices.WorkingDatabase)
            using (var namedObjectDict = (DBDictionary)transaction.GetObject(
                    db.NamedObjectsDictionaryId,
                    OpenMode.ForRead))
            {
                if (!namedObjectDict.Contains(dictionaryName))
                {
                    throw new Exception(
                        $"The dictionary named '{dictionaryName}' does not exist in this DWG file");
                }

                var dictionary = transaction.GetObject(
                    namedObjectDict.GetAt(dictionaryName),
                    OpenMode.ForWrite) as DBDictionary;

                if (!dictionary.Contains(key))
                {
                    dictionary.
                }

                var xRecord = (Xrecord)transaction.GetObject(
                    dictionary.GetAt(key),
                    OpenMode.ForRead);

                //string value = string.Empty;
                foreach (TypedValue typedValue in xRecord.Data.AsArray())
                {
                    if (typedValue.TypeCode == 1)
                    {
                        value = typedValue.Value as string;
                    }
                }
            }
        }*/

        public static bool HasDictionaryNamed(string dictionaryName)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var db = HostApplicationServices.WorkingDatabase)
            using (var namedObjectDict = (DBDictionary)transaction.GetObject(
                    db.NamedObjectsDictionaryId,
                    OpenMode.ForRead))
            {
                return namedObjectDict.Contains(dictionaryName);
            }
        }
    }
}
