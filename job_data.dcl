//DCL CODING STARTS HERE
job_data
 
: dialog
 
{
label = "Test Dialog No 1";
 
	: text
	{
	label = "This is a Test Message";
	alignment = centered;
	}
	
	: edit_box {
	action = "(texted)";
	allow_accept = true;
	edit_limit = 31;
	key = "EB";
	label = "File name: ";
	mnemonic = "F";
	value = "drawing"; 
	width = 30;
	}     

	: button
	{
	key = "accept";
	label = "Close";
	is_default = true;
	fixed_width = true;
	alignment = centered;
	}
	
	: button
	{
	key = "what";
	label = "What";
	//is_default = true;
	//fixed_width = true;
	alignment = centered;
	}
 
}