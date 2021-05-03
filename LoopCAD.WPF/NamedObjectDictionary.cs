using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class NamedObjectDictionary
    {
        public static string KeyValue(string dictionaryName, string key)
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
                    throw new Exception(
                        $"Key '{key}' does not exist in the dictionary '{dictionaryName}'");
                }

                var xRecord = (Xrecord)transaction.GetObject(
                    dictionary.GetAt(key),
                    OpenMode.ForRead);

                string value = string.Empty;
                foreach (TypedValue typedValue in xRecord.Data.AsArray())
                {
                    if (typedValue.TypeCode == 1)
                    {
                        value = typedValue.Value as string;
                    }
                }

                return value;
            }
        }
    }
}
