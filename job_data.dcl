//DCL CODING STARTS HERE
job_data : dialog
{
    label = "Job Data";
    
    : popup_list 
    {
        edit_width = 0;
        key = "calculated_by_company";
        label = "Calculated by Company: ";
        list = "X-Fire\n13dpex.com\nOther";
        value = "Other";
        width = 60;
        fixed_width = true;
        is_tab_stop = true;
        action = "(set-job-data $key $value)";    
    }

    : boxed_column
    {
        label = "Job";

        : edit_box 
        {
            allow_accept = true;
            edit_limit = 31;
            key = "job_number";
            label = "Job Number: ";
            mnemonic = "J";
            value = ""; 
            width = 30;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true; 
            action = "(set-job-data $key $value)";    
        }     

        : edit_box 
        {
            allow_accept = true;
            edit_limit = 31;
            key = "job_name";
            label = "Job Name: ";
            mnemonic = "N";
            value = ""; 
            width = 60;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true;     
            action = "(set-job-data $key $value)";
        }
        
        : edit_box 
        {
            allow_accept = true;
            edit_limit = 255;
            key = "job_site_address";
            label = "Site Address: ";
            value = ""; 
            width = 50;
            alignment = right;
            is_tab_stop = true; 
            action = "(set-job-data $key $value)";
        }
    }
            
    : popup_list 
    {
        key = "sprinkler_pipe_type";
        label = "Sprinkler Pipe Type: ";
        list = "Rehau PEX\nCopper\nCPVC\nOther";
        value = "Other";
        width = 50;
        fixed_width = true;
        is_tab_stop = true;
        action = "(set-job-data $key $value)";    
    }
    
    : popup_list 
    {
        key = "sprinkler_fitting_type";
        label = "Sprinkler Fitting Type: ";
        list = "Rehau Brass\nRehau Plastic\nOther";
        value = "Other";
        width = 50;
        fixed_width = true;
        is_tab_stop = true;
        action = "(set-job-data $key $value)";    
    }
        
    : boxed_row
    {
        label = "Water Supply Node";
        
        : column
        {
            : edit_box 
            {
                allow_accept = true;
                key = "supply_name";
                label = "Node Name: ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_elevation";
                label = "Elevation (ft): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;
                action = "(set-job-data $key $value)";        
            }
                
            : edit_box 
            {
                allow_accept = true;
                key = "supply_available_flow";
                label = "Available Flow (gpm): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;
                action = "(set-job-data $key $value)";    
            }
        }
        
        : column
        {
            : edit_box 
            {
                allow_accept = true;
                key = "supply_static_pressure";
                label = "Static Pressure (psi): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;
                action = "(set-job-data $key $value)";        
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_residual_pressure";
                label = "Residual Pressure (psi): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;
                action = "(set-job-data $key $value)";        
            }
        }
    }
    
    : edit_box 
    {
        allow_accept = true;
        key = "domestic_flow_added";
        label = "Domestic Flow Added (gpm): ";
        value = ""; 
        width = 50;
        fixed_width = true;
        alignment = left;
        is_tab_stop = true;
        action = "(set-job-data $key $value)";        
    }

    : boxed_row
    {
        label = "Supply to Manifold Pipe";
        
        : column 
        {
            alignment = top;
            
            : popup_list 
            {
                key = "supply_pipe_type";
                label = "Pipe Type: ";
                list = "Rehau PEX\nSpears Flameguard CPVPC\nCopper\nCPVC\nPoly\nOther";
                value = "Other";
                width = 40;
                fixed_width = true;
                is_tab_stop = true;
                action = "(set-job-data $key $value)";    
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_length";
                label = "Pipe Length (ft): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_size";
                label = "Pipe Size (inches): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_internal_diameter";
                label = "Internal Diameter (inches): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
        }
        
        : column
        {
            alignment = top;
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_fittings_summary";
                label = "Fittings Summary: ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_fittings_equiv_length";
                label = "Fittings Equiv. Length (ft): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
            
            : edit_box 
            {
                allow_accept = true;
                key = "supply_pipe_add_pressure_loss";
                label = "Add Pressure Loss (psi): ";
                value = ""; 
                width = 40;
                fixed_width = true;
                alignment = left;
                is_tab_stop = true;         
                action = "(set-job-data $key $value)";
            }
        }
    }
    
    : boxed_column
    {
        label = "Water Flow Switch";
        
        : edit_box 
        {
            allow_accept = true;
            key = "water_flow_switch_make_model";
            label = "Make && Model: ";
            value = ""; 
            width = 40;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true;         
            action = "(set-job-data $key $value)";
        }
        
        : edit_box 
        {
            allow_accept = true;
            key = "water_flow_switch_pressure_loss";
            label = "Pressure Loss (psi): ";
            value = ""; 
            width = 40;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true;         
            action = "(set-job-data $key $value)";
        }
    }
    
    : boxed_column
    {
        label = "Head Model Defaults";
        
        : edit_box 
        {
            allow_accept = true;
            key = "head_model_default";
            label = "Head Model Default: ";
            value = ""; 
            width = 40;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true;
            action = "(set-job-data $key $value)";
        }
        
        : edit_box 
        {
            allow_accept = true;
            key = "head_coverage_default";
            label = "Head Coverage Default (ft): ";
            value = ""; 
            width = 40;
            fixed_width = true;
            alignment = left;
            is_tab_stop = true;
            action = "(set-job-data $key $value)";
        }
    }
    
    allow_accept = true;
    ok_cancel;
}