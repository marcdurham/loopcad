using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace LoopCAD.WPF
{
    // This Jig will show a block as a big mouse cursor
    public class BlockJig : DrawJig
    {
        public Point3d point;
        private ObjectId blockDefId = ObjectId.Null;

        // Shows the block until the user clicks a point.
        public PromptResult DragMe(ObjectId blockDefId, out Point3d point)
        {
            this.blockDefId = blockDefId;
            PromptResult result = Editor().Drag(this);
            point = this.point;
            
            return result;
        }

        // Update position of the block.
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var options = new JigPromptPointOptions()
            {
                UserInputControls = (UserInputControls.Accept3dCoordinates |
                    UserInputControls.NullResponseAccepted)
            };

            options.Message = "Select a point:";
            PromptPointResult result = prompts.AcquirePoint(options);

            Point3d currentPoint = result.Value;
            if (currentPoint == point)
            {
                return SamplerStatus.NoChange;
            }

            point = currentPoint;
            if (result.Status == PromptStatus.OK)
            {
                return SamplerStatus.OK;
            }

            return SamplerStatus.Cancel;
        }

        // Show block in its current position
        protected override bool WorldDraw(WorldDraw draw)
        {
            var inMemoryBlockRef = new BlockReference(point, blockDefId);
            draw.Geometry.Draw(inMemoryBlockRef);
            inMemoryBlockRef.Dispose();

            return true;
        }

        static Editor Editor()
        {
            return Application.DocumentManager.MdiActiveDocument.Editor;
        }
    } 
}
