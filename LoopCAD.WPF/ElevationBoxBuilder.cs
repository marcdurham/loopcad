using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LoopCAD.WPF
{
    public class ElevationBoxBuilder
    {
        private const string LayerName = "ElevationBox";

        public static void Start(Point3d firstCorner, Point3d otherCorner, double elevation)
        {

            using (var t = ModelSpace.StartTransaction())
            {
                ObjectId layerId = Layer.Ensure(LayerName, ColorIndices.Magenta);

                var rectangle = new Polyline(3)
                {
                    Layer = LayerName,
                    Closed = true,
                    ColorIndex = ColorIndices.ByLayer,
                };

                rectangle.AddVertexAt(0, new Point2d(firstCorner.X, firstCorner.Y), 0, 0, 0);
                rectangle.AddVertexAt(1, new Point2d(otherCorner.X, firstCorner.Y), 0, 0, 0);
                rectangle.AddVertexAt(2, new Point2d(otherCorner.X, otherCorner.Y), 0, 0, 0);
                rectangle.AddVertexAt(3, new Point2d(firstCorner.X, otherCorner.Y), 0, 0, 0);

                ModelSpace.From(t).AppendEntity(rectangle);
                t.AddNewlyCreatedDBObject(rectangle, true);

                var mText = new MText()
                {
                    Contents = $"Elevation {elevation}",
                    Location = firstCorner,
                    Layer = LayerName,
                    TextHeight = 10.0,
                    ColorIndex = ColorIndices.ByLayer,
                };

                ModelSpace.From(t).AppendEntity(mText);
                t.AddNewlyCreatedDBObject(mText, true);

                t.Commit();
            }
        }
    }
}
