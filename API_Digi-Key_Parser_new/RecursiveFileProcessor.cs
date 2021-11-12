using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_Digi_Key_Parser_new
{
    public class RecursiveFileProcessor
    {
        private string[] getPath;
        private List<string> outPath = new List<string>();
        public string[] GetPath
        {
            get
            {
                return getPath;
            }
            private set
            {
                getPath = value;
            }
        }
        public List<string> OutPath
        {
            get
            {
                return outPath;
            }
            private set
            {
                outPath = value;
            }
        }
        public RecursiveFileProcessor(string[] GetPath)
        {
            this.getPath = GetPath;
        }
        public void RunProcessor(string[] path)
        {
            foreach (string p in path)
            {
                if (File.Exists(p))
                {
                    // This path is a file
                    ProcessFile(p);
                }
                else if (Directory.Exists(p))
                {
                    // This path is a directory
                    ProcessDirectory(p);
                }
                else
                {
                    //return $"{p} is not a valid file or directory.";
                }
            }
        }
        public void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }
        public void ProcessFile(string path)
        {
            OutPath.Add(path);
        }
    }
}
