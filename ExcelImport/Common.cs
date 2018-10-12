using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace ExcelImport
{
    public static class Common
    {
        public static void WriteToFile(string strLineToWright, bool isFirstLine)
        {
            string rootPath = AppDomain.CurrentDomain.BaseDirectory;
            string strPath = rootPath+"\\LogFile.txt";
            if (isFirstLine)
            {
                System.IO.File.WriteAllText(strPath, " Log Date and Time: " + DateTime.Now.ToString() + " Log content: " + strLineToWright);
            }
            else
            {
                using (System.IO.StreamWriter file =
               new System.IO.StreamWriter(strPath, true))
                { 
                   file.WriteLine(" Log Date and Time: " + DateTime.Now.ToString() + " Log content: " + strLineToWright);
                }
            }
        }
    }
}