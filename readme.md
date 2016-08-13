# Word macros for everyday use
## Manage Xml
### Add Library reference
1. Choose Tools > Reference
2. Pick ```Microsft XML, v6.0``` (Or whatever version. But mostly the latest. Unless you've got some reason)

Sample: ```xml.bas```

## Manage files with FileSystemObject
### Includes
* Folder object and iterating files in folder
* File object and file attributes
* Read and write files

1. Choose Tools > Reference
2. Pick Microsoft Scripting Runtime

Sample: ```scripting.bas```

### File Attributes
<table>
<tr><td>Constant</td><td>Value</td><td>Description</td></tr>
<tr><td>Normal</td><td>0</td><td>Normal file. No attributes are set.</td></tr>
<tr><td>ReadOnly</td><td>1</td><td>Read-only file. Attribute is read/write.</td></tr>
<tr><td>Hidden</td><td>2</td><td>Hidden file. Attribute is read/write.</td></tr>
<tr><td>System</td><td>4</td><td>System file. Attribute is read/write.</td></tr>
<tr><td>Volume</td><td>8</td><td>Disk drive volume label. Attribute is read-only.</td></tr>
<tr><td>Directory</td><td>16</td><td>Folder or directory. Attribute is read-only.</td></tr>
<tr><td>Archive</td><td>32</td><td>File has changed since last backup. Attribute is read/write.</td></tr>
<tr><td>Alias</td><td>1024</td><td>Link or shortcut. Attribute is read-only.</td></tr>
<tr><td>Compressed</td><td>2048</td><td>Compressed file. Attribute is read-only.</td></tr>
</table>


## Iterate tables in Word doc
### Includes
* Iterate tables in a doc
* Iterate rows in a table
* Iterate cells in a row
* Get the text from a cell

Sample: ```tables.bas```