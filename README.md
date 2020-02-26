# Excel_Automation
This is a simple code to automate manual excel update work. There is a mapping sheet tagged along, in which User will have to specify the details of source and target file, sheet, rows in scope and column mapping.
This is what the code will do for each row in the source file: It will take the first column(supposed to be distinguishable key ID) and search for the same value in target file's first column. If it finds a match,  it will update all the columns specified in the mapping for that row accordingly.
