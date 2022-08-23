using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace CNO.BPA.GenerateXLS.Framework
{
    class Cryptography
    {
        static string passPhrase = "OC59gg1";
        static string saltValue = "val@%^ff44";        // can be any string, used in password generation
        static string hashAlgorithm = "MD5";             // can be "MD5" or "SHA1", algorithm used in password generation
        static int passwordIterations = 2;                  // can be any number, number of iterations used in password generation
        static string initVector = "5ki2c3rr3AAh87oo"; // must be 16 bytes, required to encrypt the first block of plaintext
        static int keySize = 192;                // can be 192 or 128 or 256, size of encryption key

        /// <summary>
        /// Encrypts specified plaintext using Rijndael symmetric 
        /// key algorithm and returns a base64-encoded result.
        /// </summary>
        /// <param name="plainText"></param>
        /// <returns>encrypted string</returns>  
        /// <remarks>changes:   DAT   09/08/2006  created</remarks>        
        public string Encrypt(string plainText)
        {
            try
            {

                // Convert strings into byte arrays.
                // Let us assume that strings only contain ASCII codes.
                // If strings include Unicode characters, use Unicode, UTF7, or UTF8 
                // encoding.
                byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
                byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

                // Convert our plaintext into a byte array.
                // Let us assume that plaintext contains UTF8-encoded characters.
                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

                // First, we must create a password, from which the key will be derived.
                // This password will be generated from the specified passphrase and 
                // salt value. The password will be created using the specified hash 
                // algorithm. Password creation can be done in several iterations.
                PasswordDeriveBytes password = new PasswordDeriveBytes(
                                                                passPhrase,
                                                                saltValueBytes,
                                                                hashAlgorithm,
                                                                passwordIterations);

                // Use the password to generate pseudo-random bytes for the encryption
                // key. Specify the size of the key in bytes (instead of bits).
                byte[] keyBytes = password.GetBytes(keySize / 8);

                // Create uninitialized Rijndael encryption object.
                RijndaelManaged symmetricKey = new RijndaelManaged();

                // It is reasonable to set encryption mode to Cipher Block Chaining
                // (CBC). Use default options for other symmetric key parameters.
                symmetricKey.Mode = CipherMode.CBC;

                // Generate encryptor from the existing key bytes and initialization 
                // vector. Key size will be defined based on the number of the key 
                // bytes.
                ICryptoTransform encryptor = symmetricKey.CreateEncryptor(
                                                                 keyBytes,
                                                                 initVectorBytes);

                // Define memory stream which will be used to hold encrypted data.
                MemoryStream memoryStream = new MemoryStream();

                // Define cryptographic stream (always use Write mode for encryption).
                CryptoStream cryptoStream = new CryptoStream(memoryStream,
                                                             encryptor,
                                                             CryptoStreamMode.Write);
                // Start encrypting.
                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);

                // Finish encrypting.
                cryptoStream.FlushFinalBlock();

                // Convert our encrypted data from a memory stream into a byte array.
                byte[] cipherTextBytes = memoryStream.ToArray();

                // Close both streams.
                memoryStream.Close();
                cryptoStream.Close();

                // Convert encrypted data into a base64-encoded string.
                string cipherText = Convert.ToBase64String(cipherTextBytes);

                // Return encrypted string.
                return cipherText;

            }
            catch (Exception ex)
            {
                throw new Exception("Cryptography.Encrypt: " + ex.Message);
            }

        }
        /// <summary>
        /// Decrypts specified ciphertext using Rijndael 
        /// symmetric key algorithm.
        /// </summary>
        /// <param name="cipherText"></param>
        /// <returns>decrypted string</returns>  
        /// <remarks>changes:   DAT   09/08/2006  created</remarks>
        public string Decrypt(string cipherText)
        {
            try
            {
                // Convert strings defining encryption key characteristics into byte
                // arrays. Let us assume that strings only contain ASCII codes.
                // If strings include Unicode characters, use Unicode, UTF7, or UTF8
                // encoding.
                byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
                byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

                // Convert our ciphertext into a byte array.
                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);

                // First, we must create a password, from which the key will be 
                // derived. This password will be generated from the specified 
                // passphrase and salt value. The password will be created using
                // the specified hash algorithm. Password creation can be done in
                // several iterations.
                PasswordDeriveBytes password = new PasswordDeriveBytes(
                                                                passPhrase,
                                                                saltValueBytes,
                                                                hashAlgorithm,
                                                                passwordIterations);

                // Use the password to generate pseudo-random bytes for the encryption
                // key. Specify the size of the key in bytes (instead of bits).
                byte[] keyBytes = password.GetBytes(keySize / 8);

                // Create uninitialized Rijndael encryption object.
                RijndaelManaged symmetricKey = new RijndaelManaged();

                // It is reasonable to set encryption mode to Cipher Block Chaining
                // (CBC). Use default options for other symmetric key parameters.
                symmetricKey.Mode = CipherMode.CBC;

                // Generate decryptor from the existing key bytes and initialization 
                // vector. Key size will be defined based on the number of the key 
                // bytes.
                ICryptoTransform decryptor = symmetricKey.CreateDecryptor(
                                                                 keyBytes,
                                                                 initVectorBytes);

                // Define memory stream which will be used to hold encrypted data.
                MemoryStream memoryStream = new MemoryStream(cipherTextBytes);

                // Define cryptographic stream (always use Read mode for encryption).
                CryptoStream cryptoStream = new CryptoStream(memoryStream,
                                                              decryptor,
                                                              CryptoStreamMode.Read);

                // Since at this point we don't know what the size of decrypted data
                // will be, allocate the buffer long enough to hold ciphertext;
                // plaintext is never longer than ciphertext.
                byte[] plainTextBytes = new byte[cipherTextBytes.Length];

                // Start decrypting.
                int decryptedByteCount = cryptoStream.Read(plainTextBytes,
                                                           0,
                                                           plainTextBytes.Length);

                // Close both streams.
                memoryStream.Close();
                cryptoStream.Close();

                // Convert decrypted data into a string. 
                // Let us assume that the original plaintext string was UTF8-encoded.
                string plainText = Encoding.UTF8.GetString(plainTextBytes,
                                                           0,
                                                           decryptedByteCount);

                // Return decrypted string.   
                return plainText;
            }
            catch (Exception ex)
            {
                throw new Exception("Cryptography.Decrypt: " + ex.Message);
            }

        }
        /// <summary>
        /// Encrypts string value to Base64 Encoding
        /// </summary>
        /// <param name="toEncode"></param>
        /// <returns>encoded string</returns>
        /// <remarks>changes:   BEH   03/19/2009  created</remarks>
        public string EncodeTo64(string toEncode)
        {
            try
            {
                byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII
                   .GetBytes(toEncode);
                string encodedValue = System.Convert.ToBase64String
                   (toEncodeAsBytes);
                return encodedValue;
            }
            catch (Exception ex)
            {
                throw new Exception("Cryptography.EncodeTo64: " + ex.Message);
            }
        }
        /// <summary>
        /// Decrypts string value from Base64 Encoding
        /// </summary>
        /// <param name="toEncode"></param>
        /// <returns>encoded string</returns>
        /// <remarks>changes:   BEH   03/19/2009  created</remarks>
        public string DecodeFrom64(string encodedData)
        {
            try
            {
                byte[] encodedDataAsBytes = System.Convert.FromBase64String
                   (encodedData);
                string decodedValue = System.Text.ASCIIEncoding.ASCII.GetString
                   (encodedDataAsBytes);
                return decodedValue;
            }
            catch (Exception ex)
            {
                throw new Exception("Cryptography.DecodeFrom64: " + ex.Message);
            }
        }
    }
}
