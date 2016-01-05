
NOTE: this program is very old. it is programmed poorly, and can be dangerous to your file system!
it is however still handy...USE AT YOUR OWN RISK!!!


this program represents my wish list of macros I wish explorer had built in...

the framework provided has implements a simple ruleset to either include or exclude certain files found based on there extension.

drag and drop a folder in the main pane to add it to the
processing list...if you add a folder all files meeting your ruleset
will be processed...if you set the recursive check box then all action will be taken
on all files in every subfolder below the one you dropped in..be carefule with this!
the frame work is easily extensible to whatever you may wish to add.

to use the Zipping features you will need to upgrade to winzip8 and download the command line addon's..(search winzip readme index for command-line to get url) you will also need to put a copy of the 2 addon exe's somewhere you have your path variable set (like c:\windows)


if you want to addany features it is pretty simple...just

1) add another element to option1 array

2) add to select statement in option1.click for label and default ruleset

3) module1 add to FolderEngine and FileEngine Selectcase statements

4) follow format of creating actual subs (just include the top if statement) the rest should be taken care of)



and not a word about no option explicit please thank you. :P