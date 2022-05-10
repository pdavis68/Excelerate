using System.IO;
using System.Linq;
using System.IO.Compression;
using Xunit;

namespace Excelerate.Test
{
    public class TempFolderTester
    {
        
        [Fact]
        public void TestZipFileCreate()
        {
            string path = string.Empty;
            using (var tf = new TempFolder())
            {
                path = tf.FolderPath;
                File.WriteAllText(Path.Combine(path, "MyFile.txt"), "This is some text");
                Directory.CreateDirectory(Path.Combine(path, "Folder1"));
                File.WriteAllText(Path.Combine(path, "Folder1", "FileInFolder1.txt"), "This is some text");
                var data = tf.CreateZipFile();
                File.WriteAllBytes(@"C:\temp\temp.zip", data);
            }

            // Make sure the zip file looks how we expect.
            Assert.True(File.Exists(@"C:\temp\temp.zip"));
            var za = ZipFile.OpenRead(@"C:\temp\temp.zip");
            Assert.Contains(za.Entries, p=>p.Name == "MyFile.txt");
            Assert.Contains(za.Entries, p=>p.Name == "FileInFolder1.txt");
            Assert.Equal("Folder1/FileInFolder1.txt", za.Entries.First(p=>p.Name == "FileInFolder1.txt").FullName);
            // Ensure directory deleted
            Assert.False(Directory.Exists(path));
        }
    }
}