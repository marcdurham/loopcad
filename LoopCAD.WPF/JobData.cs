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
        public string SupplyStaticPressure { get; set; }
        public string SupplyResidualPressure { get; set; }
        public string SupplyAvailableFlow { get; set; }
        public string SupplyElevation { get; set; }
        public string SupplyPipeType { get; set; }
        public string SupplyPipeSize { get; set; }
        public string SupplyPipeInternalDiameter { get; set; }
        public string SupplyPipeCFactor { get; set; }
        public string SupplyPipeLength { get; set; }
        public string SupplyName { get; set; }
        public string DomesticFlowAdded { get; set; }
        public string WaterFlowSwitchMakeModel { get; set; }
        public string WaterFlowSwitchPressureLoss { get; set; }
        public string SupplyPipeFittingsSummary { get; set; }
        public string SupplyPipeFittingsEquivLength { get; set; }
        public string SupplyPipeAddPressureLoss { get; set; }
        public string HeadModelDefault { get; set; }
        public string HeadCoverageDefault { get; set; }

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

        void GetValues()
        {
            var form = new JobDataForm();
            form.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            form.ShowDialog();

            var properties = typeof(JobData).GetProperties();
            foreach(var property in properties)
            {
                string key = SnakeCase.Convert(property.Name);
                string value = NamedObjectDictionary.KeyValue("job_data", key);
                property.SetValue(this, value);
            }
        }

        void SetValues()
        {
            var properties = typeof(JobData).GetProperties();
            foreach (var property in properties)
            {
                string key = SnakeCase.Convert(property.Name);
                string value = NamedObjectDictionary.KeyValue("job_data", key);
                property.SetValue(this, value);
            }
        }
    }
}
