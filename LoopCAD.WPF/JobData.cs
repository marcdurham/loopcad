using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;

namespace LoopCAD.WPF
{
    public class JobData
    {
        public bool HasJobDataDictionary { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
        public string JobSiteAddress { get; set; }
        public string CalculatedByCompany { get; set; }
        public string SprinklerPipeType { get; set; }
        public string SprinklerFittingType { get; set; }
        public string SupplyName { get; set; }
        public string SupplyStaticPressure { get; set; }
        public string SupplyResidualPressure { get; set; }
        public string SupplyAvailableFlow { get; set; }
        public string SupplyElevation { get; set; }
        public string SupplyPipeType { get; set; }
        public string SupplyPipeSize { get; set; }
        public string SupplyPipeInternalDiameter { get; set; }
        public string SupplyPipeCFactor { get; set; }
        public string SupplyPipeLength { get; set; }
        public string SupplyPipeFittingsSummary { get; set; }
        public string SupplyPipeFittingsEquivLength { get; set; }
        public string SupplyPipeAddPressureLoss { get; set; }
        public string DomesticFlowAdded { get; set; }
        public string WaterFlowSwitchMakeModel { get; set; }
        public string WaterFlowSwitchPressureLoss { get; set; }
        public string HeadModelDefault { get; set; }
        public string HeadCoverageDefault { get; set; }
        public List<string> CalculatedByCompanies { get; set; }
        public List<string> SprinklerPipeTypes{ get; set; }
        public List<string> SprinklerFittingTypes{ get; set; }
        public List<string> SupplyPipeTypes{ get; set; }

        public static JobData Load()
        {
            var data = new JobData();

            data.HasJobDataDictionary = NamedObjectDictionary
                .HasDictionaryNamed("job_data");

            if (data.HasJobDataDictionary)
            {
                data.GetValues();
            }
            else
            {
                data = JobDataBlock.Load();
            }

            return data;
        } 
        
        public void Save()
        {
            HasJobDataDictionary = NamedObjectDictionary
                .HasDictionaryNamed("job_data");

            if (HasJobDataDictionary)
            {
                SetValues();
            }
        }

        void GetValues()
        {
            var properties = typeof(JobData).GetProperties();
            foreach(var property in properties)
            {
                if (property.Name == nameof(HasJobDataDictionary))
                {
                    continue;
                }

                string key = SnakeCase.Convert(property.Name);
                if (NamedObjectDictionary.KeyValue("job_data", key, out string value))
                {
                    property.SetValue(this, value);
                }
            }
        }

        void SetValues()
        {
            var properties = typeof(JobData).GetProperties();
            foreach (var property in properties)
            {
                string key = SnakeCase.Convert(property.Name);
                if (NamedObjectDictionary.KeyValue("job_data", key, out string _))
                {
                    string v = property.GetValue(this) as string;
                    NamedObjectDictionary.SetKeyValue("job_data", key, new ResultBuffer(new TypedValue(typeCode: 1, value: v)));
                }
            }
        }
    }
}
