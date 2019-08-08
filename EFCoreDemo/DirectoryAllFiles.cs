using System;
using System.Collections.Generic;
using System.IO;

namespace EFCoreDemo
{
    public class DirectoryAllFiles
    {
        static List<FileInfo> FileList = new List<FileInfo>();
        public static List<FileInfo> GetAllFiles(string dir)
        {
            DirectoryInfo di = new DirectoryInfo(dir);
            FileInfo[] allFile = di.GetFiles();
            foreach (FileInfo fi in allFile)
            {
                FileList.Add(fi);
            }
            DirectoryInfo[] allDir = di.GetDirectories();
            foreach (DirectoryInfo d in allDir)
            {
                GetAllFiles(d.FullName);
            }
            return FileList;
        }
    }    
}
