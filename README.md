# FolderRegister
## Excel spreadsheet showing folders and subfolders

### 1. Introduction

This repository has the objective of developing a spreadsheet in Microsoft Excel showing the folders and subfolders that exist at the location where the spreadsheet is saved. For this application VBA with Microsoft Excel is used with the objective of making the comprehension and usage of it easier to co-workers and colleagues accustomed to work with Microsoft Excel.

### 2. Development

One of the requirements for this application to work is the update of the folders and subfolders showned on the spreadsheet when changing (cut/copy & paste) the Excel file from one location to another location. If not, the user would have to manually run the Excel file every time a new copy was pasted.

The folders and subfolders names are showned in a column hierarchy basis, where the first column corresponds to the folder in which the file is saved; the second column corresponds to folders saved inside the main folder (these folders will be at the same directory as the Excel file); and so on. Note also that each folder and subfolder is displayed in a separate line. The overall display of the folders on the spreadsheet is of a tree diagram type.

Additionally, the application displays for each filled column a different gradient of green background colour, in order to distinguish folders hierarchy, and each folder name when clicked opens a window showing the contents of the correspondent folder.

### 3. Instructions

1. Copy and paste the file at a certain directory;
2. Open the file;
3. The file shows a hierarchy of folders starting from the folder where the file is saved.
4. Click on a folder's name to go directly into the folder
