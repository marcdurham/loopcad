using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;
using System.Diagnostics;

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

            return data ?? new JobData();
        } 
        
        public void Save()
        {
            HasJobDataDictionary = NamedObjectDictionary
                .HasDictionaryNamed("job_data");

            if (!HasJobDataDictionary)
            {
                NamedObjectDictionary.CreateDictionary("job_data");
            }
 
            SetValues();
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
                if (NamedObjectDictionary.KeyValue("job_data", key, out string value)
                    && (property.PropertyType == typeof(string)
                     || property.PropertyType == typeof(int)
                     || property.PropertyType == typeof(double)))
                {
                    if(property.PropertyType == typeof(string))
                        property.SetValue(this, value);
                    else if (property.PropertyType == typeof(int))
                        property.SetValue(this, int.Parse(value));
                    else if (property.PropertyType == typeof(double))
                        property.SetValue(this, double.Parse(value));

                    Debug.WriteLine($"Loading XRecord Key: {key} Value: {value} (Property: {property.Name})");
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
                    Debug.WriteLine($"Writing XRecord Key: {key} Value: {v} (Property: {property.Name})");
                    NamedObjectDictionary.SetKeyValue("job_data", key, v);
                }
            }
        }
    }
}
