using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LoopCAD.WPF
{
    public class RiserLabel
    {
        public const string LayerName = "RiserLabels";
        public const string BlockName = "RiserLabel";
        public const string TagName = "RISERNUMBER";

        public static int HighestNumber()
        {
            using (var transaction = ModelSpace.StartTransaction())
            {
                int lastNumber = 0;
                var labelIds = GetRiserLabelIds(transaction);

                foreach (var id in labelIds)
                {
                    string text = AttributeReader.TextString(transaction, id, BlockName, tag: TagName);
                    var match = Regex.Match(text, @"R\.(\d+)\.[A-Z]");
                    if (match.Success)
                    {
                        string numberString = match.Groups[1].Value;
                        int number = int.Parse(numberString);
                        if (number > lastNumber)
                        {
                            lastNumber = number;
                        }
                    }
                }

                return lastNumber;
            }
        }

        public static char HighestSuffix()
        {
            using (var transaction = ModelSpace.StartTransaction())
            {
                byte lastNumber = 'A' - 1;

                foreach (string text in GetRiserLabelTexts())
                {
                    var match = Regex.Match(text, @"R\.(\d+)\.([A-Z])");
                    if (match.Success)
                    {
                        string suffixString = match.Groups[2].Value;
                        byte number = (byte)suffixString[0];
                        if (number > lastNumber)
                        {
                            lastNumber = number;
                        }
                    }
                }

                return (char)lastNumber;
            }
        }

        static List<string> GetRiserLabelTexts()
        {
            using (var transaction = ModelSpace.StartTransaction())
            {
                var texts = new List<string>();
                var labelIds = GetRiserLabelIds(transaction);

                foreach (var id in labelIds)
                {
                    string text = AttributeReader.TextString(transaction, id, BlockName, tag: TagName);
                    texts.Add(text);
                }

                return texts;
            }
        }

        static List<ObjectId> GetRiserLabelIds(Transaction transaction)
        {
            var labels = new List<ObjectId>();
            foreach (var objectId in ModelSpace.From(transaction))
            {
                if (IsRiserLabel(transaction, objectId))
                {
                    labels.Add(objectId);
                }
            }

            return labels;
        }

        static bool IsRiserLabel(Transaction transaction, ObjectId objectId)
        {
            var block = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            return objectId.ObjectClass.DxfName == "INSERT" &&
                string.Equals(block.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(block.Name, BlockName, StringComparison.OrdinalIgnoreCase);
        }
    }
}
