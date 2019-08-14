# libBAUtil - General purpose helper methods

This library includes helper methods for all sorts of different areas. Some are obvious ports of 
similar methods I have used and developed for VB6 over the time, that made porting old applications to VB.NET 
more easy.

Please note that I _didn't_ develop _all_ of these myself. Where possible/available, the original source is mentioned.

The following classes and methods are included:

## FilesystemUtil
- DenormalizePath()    
Ensure a path does NOT end with a path delimiter.
- NormalizePath()    
Ensure a path does end with a path delimiter.
- BackupFile()    
Creates a backup of a file by copying/moving it from the source to the target folder.


## MathUtil

## ObjectUtil

## StringUtil
- Bytes2FormattedString()    
Creates a formatted string representing the size in its proper 'spelled out' unit _(Bytes, KB etc.)._
- MCase()    
Capitalize the first letter of a string.
- Left() / Right() / Mid()    
Emulate the respective VB6/VBA methods.
- Space()    
- vbNewLine, vbTab    
Replacements for various handy VB6 string constants.

## FixedLengthString
Emulates a VB6/VBA fixed length string (```MyString * 3```)
