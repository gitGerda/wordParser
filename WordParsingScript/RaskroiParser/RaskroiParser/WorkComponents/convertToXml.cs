using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace RaskroiParser.WorkComponents
{
    public static class convertToXml
    {
        public static void convert(string pathToScript, string pathToWorkDir,string pathToParsedFilesDir)
        {
            Process convertDocToXml = new Process();
            convertDocToXml.StartInfo.FileName = pathToScript;
            convertDocToXml.StartInfo.Arguments = pathToWorkDir+" "+pathToParsedFilesDir;
            //convertDocToXml.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            convertDocToXml.Start();

            convertDocToXml.WaitForExit();
            Console.WriteLine(345345);
        }
    }
}
