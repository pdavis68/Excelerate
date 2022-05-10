using System;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace Excelerate
{

    /// <summary>
    /// Manages the temp folder used to generate the .xlsx file. Zips up the data and
    /// deletes all the temp files when done.
    /// </summary>
    public class TempFolder : IDisposable
    {
        private string _path = string.Empty;

        public TempFolder()
        {
            _path = Path.GetTempFileName();
            File.Delete(_path);
            Directory.CreateDirectory(_path);
        }

        public string FolderPath { get => _path; }

        public byte[] CreateZipFile()
        {
            using (var ms = new MemoryStream())
            using (var zipStream = new ZipOutputStream(ms))
            {
                zipStream.UseZip64 = UseZip64.Off;
                RecursiveAddToZipFile(zipStream, _path);
                zipStream.Finish();
                zipStream.Close();
                return ms.ToArray();
            }
        }

        private void RecursiveAddToZipFile(ZipOutputStream zipStream, string path)
        {
            foreach(var dir in Directory.GetDirectories(path))
            {
                var entry = new ZipEntry(dir.Replace(_path, "") + @"/");
                entry.CompressionMethod = CompressionMethod.Deflated;
                RecursiveAddToZipFile(zipStream, dir);
                zipStream.PutNextEntry(entry);
                zipStream.CloseEntry();
            }

            var files = Directory.GetFiles(path);
            foreach(var file in files)
            {
                var entry = new ZipEntry(file.Replace(_path, ""));
                entry.DateTime = DateTime.Now;
                entry.CompressionMethod = CompressionMethod.Deflated;
                zipStream.PutNextEntry(entry);
                using (FileStream fs = File.OpenRead(file)) 
                {
                    int totalBytes = 0;
                    byte[] buffer = new byte[1024 * 32];
                    int bytesRead;
                    do
                    {
                        bytesRead = fs.Read(buffer, 0, buffer.Length);
                        totalBytes += bytesRead;
                        zipStream.Write(buffer, 0, bytesRead);
                    } while (bytesRead > 0);
                    entry.Size = totalBytes;
                }
                zipStream.CloseEntry();
            }
        }

        public void Dispose()
        {
            DeleteAllRecursive(_path);
            Directory.Delete(_path);
            _path = null;
        }

        private void DeleteAllRecursive(string path)
        {
            string[] directories = Directory.GetDirectories(path);
            foreach(var dir in directories)
            {
                DeleteAllRecursive(dir);
                Directory.Delete(dir);
            }
            string[] files = Directory.GetFiles(path);
            foreach(var file in files)
            {
                File.Delete(file);
            }
        }
    }
}