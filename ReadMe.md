# libBAUtil - General purpose helper methods

This library includes helper methods for all sorts of different areas. Some are obvious ports of 
similar methods I have used and developed for VB6 over the time, that made porting old applications to VB.NET 
more easy.

Please note that I _didn't_ develop _all_ of these myself. Where possible/available, the original source is mentioned.

The following classes and methods are included:

## CryptoUtil
- baCrypto3DES.DecryptData()/EnCryptData()    
String de-/encryption using 3DES.
- baCryptoAES.DecryptData()/EnCryptData()    
String de-/encryption using AES.


## FilesystemUtil
- DenormalizePath()    
Ensure a path does NOT end with a path delimiter.
- NormalizePath()    
Ensure a path does end with a path delimiter.
- BackupFile()    
Creates a backup of a file by copying/moving it from the source to the target folder.


## MathUtil
Returns the % of Total given by Part, e.g. Total = 200, Part = 50 = 25(%)


## ObjectUtil
- GetEnumNameFromValue    
Returns the Enumeration member's name for the specific value.
- IsSerializable    
Determine if an object is serializable.
- Serialize    
Serialze an object to (a) XML (string).
- Deserialize    
Deserialize a XML (string) to an object.
- DeserializeAsClass    
Deserialize a XML (string) to a specific class.


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
Emulates a VB6/VBA fixed length string (e.g. ```MyString * 3```).
