Imports System.Security.Cryptography

''' <summary>
''' En-/Decryption helper (3DES).
''' </summary>
''' <remarks>
''' Source: https://msdn.microsoft.com/en-us/library/ms172831(v=vs.110).aspx
''' </remarks>
Public NotInheritable Class baCrypto3DES

#Region "Declares"

   Private TripleDes As New TripleDESCryptoServiceProvider

#End Region

#Region "Methods - Public"

   ''' <summary>
   ''' Decode a string.
   ''' </summary>
   ''' <param name="encryptedText">Encoded string</param>
   ''' <returns>Decoded <paramref name="encryptedText"/></returns>
   Public Function DecryptData(ByVal encryptedText As String) As String

      ' Convert the encrypted text string to a byte array. 
      Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedText)

      ' Create the stream. 
      Dim ms As New System.IO.MemoryStream
      ' Create the decoder to write to the stream. 
      Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

      ' Use the crypto stream to write the byte array to the stream.
      decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
      decStream.FlushFinalBlock()

      ' Convert the plaintext stream to a string. 
      Return System.Text.Encoding.Unicode.GetString(ms.ToArray)

   End Function

   ''' <summary>
   ''' Encode a string.
   ''' </summary>
   ''' <param name="plainText">Plain string</param>
   ''' <returns>Encoded <paramref name="plainText"/></returns>
   Public Function EncryptData(ByVal plainText As String) As String

      If String.IsNullOrEmpty(plainText) = True Then
         plainText = ""
      End If

      ' Convert the plaintext string to a byte array. 
      Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plainText)

      ' Create the stream. 
      Dim ms As New System.IO.MemoryStream
      ' Create the encoder to write to the stream. 
      Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

      ' Use the crypto stream to write the byte array to the stream.
      encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
      encStream.FlushFinalBlock()

      ' Convert the encrypted stream to a printable string. 
      Return Convert.ToBase64String(ms.ToArray)

   End Function

   ''' <summary>
   ''' Initializes a new instance of the CryptoUtil class
   ''' </summary>
   ''' <param name="key">De-/Encryption key ("password")</param>
   Sub New(ByVal key As String)
      ' Initialize the crypto provider.
      TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
      'TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
      TripleDes.IV = TruncateHash(String.Empty, TripleDes.BlockSize \ 8)
   End Sub

#End Region

   Private Function TruncateHash(ByVal key As String, ByVal length As Int32) As Byte()

      Dim sha1 As New SHA1CryptoServiceProvider

      ' Hash the key. 
      Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
      Dim hash() As Byte = sha1.ComputeHash(keyBytes)

      ' Truncate or pad the hash. 
      ReDim Preserve hash(length - 1)
      Return hash

   End Function

End Class   '// baCrypto3DES

''' <summary>
''' En-/Decrypting helper (AES).
''' </summary>
''' <remarks>
''' Source: https://msdn.microsoft.com/en-us/library/ms172831(v=vs.110).aspx
''' </remarks>
Public NotInheritable Class baCryptoAES

#Region "Declares"

   Private AES As New RijndaelManaged

#End Region

#Region "Methods - Public"

   ''' <summary>
   ''' Decode a string.
   ''' </summary>
   ''' <param name="encryptedText">Encoded string</param>
   ''' <returns>Decoded <paramref name="encryptedText"/></returns>
   Public Function DecryptData(ByVal encryptedText As String) As String

      ' Convert the encrypted text string to a byte array. 
      Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedText)

      ' Create the stream. 
      Dim ms As New System.IO.MemoryStream
      ' Create the decoder to write to the stream. 
      Dim decStream As New CryptoStream(ms, AES.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

      ' Use the crypto stream to write the byte array to the stream.
      decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
      decStream.FlushFinalBlock()

      ' Convert the plain text stream to a string. 
      Return System.Text.Encoding.Unicode.GetString(ms.ToArray)

   End Function

   ''' <summary>
   ''' Encode a string.
   ''' </summary>
   ''' <param name="plainText">Plain string</param>
   ''' <returns>Encoded <paramref name="plainText"/></returns>
   Public Function EncryptData(ByVal plainText As String) As String

      If String.IsNullOrEmpty(plainText) = True Then
         plainText = ""
      End If

      ' Convert the plaintext string to a byte array. 
      Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plainText)

      ' Create the stream. 
      Dim ms As New System.IO.MemoryStream
      ' Create the encoder to write to the stream. 
      Dim encStream As New CryptoStream(ms, AES.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

      ' Use the crypto stream to write the byte array to the stream.
      encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
      encStream.FlushFinalBlock()

      ' Convert the encrypted stream to a printable string. 
      Return Convert.ToBase64String(ms.ToArray)

   End Function

   ''' <summary>
   ''' Initializes a new instance of the CryptoUtil class
   ''' </summary>
   ''' <param name="key">De-/Encryption key ("password")</param>
   Sub New(ByVal key As String)
      ' Initialize the crypto provider.
      AES.Key = TruncateHash(key, AES.KeySize \ 8)
      AES.IV = TruncateHash("", AES.BlockSize \ 8)
   End Sub

#End Region

   Private Function TruncateHash(ByVal key As String, ByVal length As Int32) As Byte()

      Dim sha1 As New SHA1CryptoServiceProvider

      ' Hash the key. 
      Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
      Dim hash() As Byte = sha1.ComputeHash(keyBytes)

      ' Truncate or pad the hash. 
      ReDim Preserve hash(length - 1)
      Return hash

   End Function

End Class
