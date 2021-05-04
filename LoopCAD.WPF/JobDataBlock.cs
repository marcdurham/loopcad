using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace LoopCAD.WPF
{
    public class JobDataBlock
    {
        public static JobData Load()
        {
            using (var trans = ModelSpace.StartTransaction())
            {
                foreach (var objectId in ModelSpace.From(trans))
                {
                    if (IsJobDataBlock(trans, objectId))
                    {
                        var data = new JobData
                        {
                            CalculatedByCompany = AttributeReader.TextString(trans, objectId, "CALCULATED_BY_COMPANY"),
                            JobNumber = AttributeReader.TextString(trans, objectId, "LEAD_NUMBER"),
                            JobName = AttributeReader.TextString(trans, objectId, "JOB_NAME"),
                            JobSiteAddress = AttributeReader.TextString(trans, objectId, "SITE_LOCATION"),
                            SupplyStaticPressure = AttributeReader.TextString(trans, objectId, "STATIC_PRESSURE"),
                            SupplyResidualPressure= AttributeReader.TextString(trans, objectId, "RESIDUAL_PRESSURE"),
                            SupplyAvailableFlow= AttributeReader.TextString(trans, objectId, "AVAILABLE_FLOW"),
                            SupplyElevation= AttributeReader.TextString(trans, objectId, "METER_ELEVATION"),
                            SupplyPipeLength= AttributeReader.TextString(trans, objectId, "METER_PIPE_LENGTH"),
                            SupplyPipeInternalDiameter= AttributeReader.TextString(trans, objectId, "METER_PIPE_INTERNAL_DIAMETER"),
                        };

                        return data;
                    }
                }
            }

            return null;
        }

        static bool IsJobDataBlock(Transaction trans, ObjectId objectId)
        {
            if (objectId.IsErased)
            {
                return false;
            }

            var block = trans.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (block == null)
            {
                return false;
            }

            return
                (string.Equals(block.Layer, "JobData", StringComparison.OrdinalIgnoreCase)
                || string.Equals(block.Layer, "Job Data", StringComparison.OrdinalIgnoreCase)) 
                && block.Name.ToUpper() == "JOBDATA";
        }
    }
}
