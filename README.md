# truEUtilities
The original VB6 Code for cleaning Pro/ENGINEER Directories. 

## How truEUtilities works:

### 1. Purging
truEUtilities follows a few basic rules in order to purge.  Read through these so you have a better understanding of what the software is doing when it purges files:
 * Files must contain a numeric extension as the last extension in order to purge.
 * Files must contain 2 and only 2 "." in the file name - there may be cases where the software can't purge some of the files that it needs to - you will have to do theses manually.
 * Recycling files will greatly increase the amount of time it takes to purge.  This is a function of Windows operating systems.

### 2. Renaming
 * Only a numeric value may be used as an extension.
 * Stripping extension is not a recommended method of file management.  Try if for your self and see which method best suits your company.
 * NOTE: When stripping the extensions off of files you must use truEUtilities to purge your directories the following times.  Pro/ENGINEER will create a new file with a ".1" extension the next time you save your work and the basic purge program that comes with Pro/ENGINEER will not handle the file that doesn't have a numeric extension.  However, truEUtilities will be able to take care of this issue for you during the renaming process if the "Strip extensions from files" is selected - not with the normal rename.  This is a performance issue.
