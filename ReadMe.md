# libBAUtil - General purpose helper methods

This library includes helper methods for all sorts of different areas. Some are obvious ports of similar methods I have used and developed for VB6 over the time, that made porting old applications to VB.NET easier.

Please note that I _didn't_ develop _all_ of these myself. Where possible/available, the original source is mentioned.

The following classes and methods are included:

## CryptoUtil

- baCrypto3DES.DecryptData()/EnCryptData()  
String de-/encryption using 3DES.
- baCryptoAES.DecryptData()/EnCryptData()  
String de-/encryption using AES.

## DateTimeUtil

Implements System.DateTime and enhanced it with additional methods/properties, which are:

- GetLastDayInMonth()  
Retrieve the last day in a month.
- IATADateLong  
Return the current DateTime property as IATA date formatted _(ddMMMyy)_ string.
- IATADateShort  
Return the current DateTime property as IATA date formatted _(ddMMM)_ string.

## FilesystemUtil

- DenormalizePath()  
Ensure a path does NOT end with a path delimiter.
- NormalizePath()  
Ensure a path does end with a path delimiter.
- BackupFile()  
Creates a backup of a file by copying/moving it from the source to the target folder.

## MathUtil

- Percent()  
Returns the % of Total given by Part, e.g. Total = 200, Part = 50 = 25(%)

## ObjectUtil

- GetEnumNameFromValue  
Returns the Enumeration member's name for the specific value.
- IsSerializable  
Determine if an object is serializable.
- Serialize  
Serialize an object to (a) XML (string).
- Deserialize  
Deserialize a XML (string) to an object.
- DeserializeAsClass  
Deserialize a XML (string) to a specific class.
- Clone()  
Deep clone an object.

## StringUtil

- Bytes2FormattedString()  
Creates a formatted string representing the size in its proper 'spelled out' unit _(Bytes, KB etc.)._
- MCase()  
Capitalize the first letter of a string.
- Left() / Right() / Mid() / String() / InStr()  
Emulate the respective VB6/VBA methods.
- Space()  
Create a string with (n) spaces.
- String()  
Create a string consisting of (n) characters.
- TrimAny()  
Trim (start and end) _any_ of the passed characters from a string.
- DateYMD()  
Return a date as a string in the form yyyymmdd[[T]hhnnss]
- EnQuote()  
Enclose a string in double quotes _(")_.
- vbNewLine, vbTab, vbQuote, vbNullString  
Replacements for various handy VB6 string constants.
- vbWhiteSpace  
A string consisting of what's considered to be _white space_.

## FixedLengthString

Emulates a VB6/VBA fixed length string (e.g. ```MyString * 3```).

---

For detailed information about this assembly's objects, see the included [CHM help file](./SandCastleHelp/Help/baUtil.chm)
