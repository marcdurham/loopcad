using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LoopCAD.WPF
{
    public class ElevationBox
    {
        public const string LayerName = "ElevationBox";

        public ObjectId Id { get; set; }
        public MText Label { get; set; }
        public Polyline Box { get; set; }

        public static List<ElevationBox> GetElevationBoxes(Transaction transaction)
        {
            var boxes = new List<ObjectId>();
            var labels = new List<ObjectId>();
            foreach (var objectId in ModelSpace.From(transaction))
            {
                if (IsElevationBoxPolyline(transaction, objectId))
                {
                    boxes.Add(objectId);
                }
                else if(IsElevationBoxLabel(transaction, objectId))
                {
                    labels.Add(objectId);
                }
            }

            var elevationsBoxes = new List<ElevationBox>();
            foreach(var box in boxes)
            {
                var polyline = transaction.GetObject(box, OpenMode.ForRead) as Polyline;
                foreach (var label in labels)
                {
                    var mtext = transaction.GetObject(label, OpenMode.ForRead) as MText;
                    if(mtext == null)
                    {
                        continue;
                    }

                    for (int i = 0; i < polyline.NumberOfVertices; i++)
                    {
                        if(PointExtensions.IsNear(mtext.Location, polyline.GetPoint3dAt(i))
                            && !elevationsBoxes.Exists(b => b.Id == polyline.Id))
                        {
                            elevationsBoxes.Add(
                                new ElevationBox
                                {
                                    Id = polyline.Id,
                                    Box = polyline,
                                    Label = mtext
                                });

                            continue;
                        }
                    }
                }
            }

            return elevationsBoxes;
        }

        static bool IsElevationBoxPolyline(Transaction transaction, ObjectId objectId)
        {
            var polyline = transaction.GetObject(objectId, OpenMode.ForRead) as Polyline;
            return (objectId.ObjectClass.DxfName == "POLYLINE" || 
                    objectId.ObjectClass.DxfName == "LWPOLYLINE") &&
                string.Equals(polyline.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                polyline.NumberOfVertices >= 4;
        }

        static bool IsElevationBoxLabel(Transaction transaction, ObjectId objectId)
        {
            var mtext = transaction.GetObject(objectId, OpenMode.ForRead) as MText;
            return objectId.ObjectClass.DxfName == "MTEXT" &&
                string.Equals(mtext.Layer, LayerName, StringComparison.OrdinalIgnoreCase) &&
                mtext.Contents != null &&
                Regex.IsMatch(mtext.Contents, @"Elevation \d+");
        }
    }
}
