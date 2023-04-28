using Autodesk.AutoCAD.ApplicationServices;
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

        public static void SetKeyValue(string dictName, string key, string value)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var db = HostApplicationServices.WorkingDatabase)
            {

                DBDictionary dbDictionary =(DBDictionary)
                    transaction.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                DBDictionary dictionary;
                if (dbDictionary.Contains(dictName))
                {
                    dictionary = (DBDictionary)transaction.GetObject(dbDictionary.GetAt(dictName), OpenMode.ForWrite);
                }
                else
                {
                    dictionary = new DBDictionary();
                    dbDictionary.UpgradeOpen();
                    dbDictionary.SetAt(dictName, dictionary);
                    transaction.AddNewlyCreatedDBObject(dictionary, true);
                }
                Xrecord xRec = new Xrecord();
                var resbuf = new ResultBuffer(new TypedValue(typeCode: 1, value: value));
                xRec.Data = resbuf;
                dictionary.SetAt(key, xRec);
                transaction.AddNewlyCreatedDBObject(xRec, true);
                transaction.Commit();
            }
        }

        public static void OLDSetKeyValue(string dictionaryName, string key, string value)
        {
            using (var transaction = ModelSpace.StartTransaction())
            using (var db = HostApplicationServices.WorkingDatabase)
            using (var namedObjectDict = (DBDictionary)transaction.GetObject(
                    db.NamedObjectsDictionaryId,
                    OpenMode.ForWrite))
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
                    throw new Exception($"Dictionary cannot write value, does not contain key: {key}");
                }

                ////////dictionary.get
                ////////var xRecord = new Xrecord();
                ////////xRecord.Data = new ResultBuffer(new TypedValue(typeCode: 1, value: "VALUE"));
                
                
                //xRecord.Data.Add(new TypedValue(typeCode: 1, value: value));
                
                ////////////var _ = dictionary.SetAt(key, xRecord);


                //var xRecord = (Xrecord)transaction.GetObject(
                //    dictionary.GetAt(key),
                //    OpenMode.ForRead);

                //foreach (TypedValue typedValue in xRecord.Data.AsArray())
                //{
                //    if (typedValue.TypeCode == 1)
                //    {
                //        value = typedValue.Value as string;
                //    }
                //}
            }
        }

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
