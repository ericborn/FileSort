# FileSort
At work I was presented with a huge number of loose files that were a part of an online database export, which need to be sorted into folders according to the staff member they belong to. These files were in various formats, pdf, jpeg, etc. and were all named with unique identifiers. I was given a spreadsheet that contained a column which these unique identifiers were listed in, along with another column that contains the name of the person who these files belong to. 

I decided to tackle the problem with Python utilizing the os, pandas and shell packages to accomplish this task. First the user provides input into the program to indicate the file path and name of their excel file. This is collected via a while loop and contains an if, elif, else statement to ensure they provide a valid xls or xlsx file. The while loop will break if the path ends in .xls or .xlsx. If it does not, a message stating an invalid file extension has been provided will be printed and the loop will return back to the start to prompt the user to provide input again. The path and filename are split into seperate variables then the excel file is parsed and turned into a dataframe with Pandas. 

Since the usernames are repeated for each item belonging to that user, I fed the names into a set to make them unique. From this unique set of names, I used the os package within a for loop to iterate through the list and create a folder with their name. I included both error checking with a try except block and logging to both the console and a log file within this loop. If the folder already exists, the except will log that it cannot be created and move on to the next name. If the folder does not exist, the loop will create the folder and log success.

Finally, after all of the folders have been created, another for loop iterates through the original dataframe and moves the files into the appropriate folder based on the username associated with the file and writes a success message to the log containing the name of the file and location it was moved to. If the file cannot be located an except block writes an error to the log and console and moves onto the next row within the dataframe.

I was able to make the program run successfully, creating over xxx new folders and moving over xxx files which would have taken a human an inconceivable amount of time

I setup a little demo Excel document and some fake files for demonstration purposes. Please use the following steps to try out the code:

1.  Extract all project files into the folder of your choice (A simple path is desireable since you will need to input the path into the     progam on a later step.
2.  Copy all files from the unsorted files folder into the base directory where the the setup.xlsx and .py file lives. (Copying not cutting is key here if you want to be able to simply delete all directories and sorted files to be able to run the program again without having to manually move all files out of the folders back to the root directory.
2.  Run the file sorter.py file in the IDE of your choice.
3.  Provide the file path along with the excel file name and extension to the Python console e.g. (C:\Users\Eric\Documents\filesort        \setup.xlsx).
3a. If you do not provide a valid xls or xlsx file extension the program will print an error and reprompt to enter the path and filename.
4.  The program should run and provide details of its operations in both the console window as well as a log.txt file that will be created in the same directory as the program was just run from.
5.  If folders for the staff names already exist, or files from the unsorted folder are not found by the program, it will print errors in both the console and log file but will not cause the program to terminate prematurely.
