using SevenZipExtractor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public static class _7zClass
    {
        /// <summary>
        /// Fonction Dezippage 
        /// </summary>
        /// <param name="sourceZip"></param>
        /// <param name="destinationPath"></param>
        /// <returns></returns>
        public static bool GetSeptZip(string sourceZip, string destinationPath)
        {

            string septzpwd = "elpAVVMzKHE4QWtSWGFGKTYrTywhS1Y5Og==";

            try
            {
                septzpwd = Const.publicBloc+DecodeFrom(septzpwd) + Const.finBloc;
                string dllPath = Tools.OfficeVersion() == "x86"?Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Resources\\7z", "x86\\7z.dll"): Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Resources\\7z", "x64\\7z.dll");
                //System.Windows.Forms.MessageBox.Show(dllPath);
                //string zipPath = @".\result.zip";
                string extractPath = destinationPath;
                extractPath = Path.GetFullPath(extractPath);

                // Ensures that the last character on the extraction path
                // is the directory separator char.
                if (!extractPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                    extractPath += Path.DirectorySeparatorChar;

                using (ArchiveFile archive = new ArchiveFile(sourceZip,dllPath))
                {
                    archive.Extract(extractPath, overwrite: true, septzpwd);
                }
                return true;
            }
            catch (System.Exception ex)
            {
                //if (Directory.Exists(destinationPath))
                //    Directory.Delete(destinationPath, true);
                Tools.LogMessage("Erreur de décompression.", ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static string DecodeFrom(string encodedData)
        {
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecode_byte = Convert.FromBase64String(encodedData);
            int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
            char[] decoded_char = new char[charCount];
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
            string result = new String(decoded_char);
            return result;
        }
    }
}
