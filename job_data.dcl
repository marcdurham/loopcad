//DCL CODING STARTS HERE
job_data
 
: dialog
 
{
	label = "Job Data";
	: popup_list 
	{
		action = "(patternlsp)";
		edit_width = 0;
		key = "company";
		list = "X-Fire\n13dpex.com\nOther";
		// value = "2";
		is_tab_stop = true; 
	}


	: edit_box 
	{
		action = "(texted)";
		allow_accept = true;
		edit_limit = 31;
		key = "job_number";
		label = "Job Number: ";
		mnemonic = "F";
		value = "drawing"; 
		width = 30;
		alignment = right;
		is_tab_stop = true; 
	}     
	
	: text
	{
		label = "Job Name";
		alignment = left;
	}

	: edit_box 
	{
		action = "(texted)";
		allow_accept = true;
		edit_limit = 31;
		key = "EB";
		label = "Job name: ";
		mnemonic = "F";
		value = "drawing"; 
		width = 30;
		alignment = right;
		is_tab_stop = true; 		
	}     

	allow_accept = true;
	ok_cancel;
 
}