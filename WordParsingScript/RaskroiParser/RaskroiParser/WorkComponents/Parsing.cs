using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using System.Xml;

namespace RaskroiParser.WorkComponents
{
    public static class Parsing
    {
        public static void parsingFunc(string pathToWorkDir, string pathToParsedFilesDir)
        {
            pathToWorkDir = @"a:\WORK\PROJECTS\WordParsingScript\TestXML\";
            /*Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            DirectoryInfo dirInfo = new DirectoryInfo(pathToWorkDir);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");

            word.Visible = false;
            word.ScreenUpdating = false;

            XmlDocument xmlDoc = new XmlDocument();

            foreach (FileInfo wordFile in wordFiles)
            {
                string fileName = wordFile.Name;
                if(!fileName.Contains("[parsed]"))
                {
                    try
                    {
                        Object filenamePath = (Object)wordFile.FullName;
                        Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filenamePath);

                        doc.Activate();
                        object outputFileName = wordFile.FullName.Replace(".doc", ".xml");
                        object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML;
                        doc.SaveAs(ref outputFileName, ref fileFormat);

                        object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                        ((Microsoft.Office.Interop.Word._Document)doc).Close(ref saveChanges);
                        doc = null;

                        xmlDoc.Load(outputFileName.ToString());

                        XmlNodeList documentBody = xmlDoc.SelectNodes("//pkg:package/" +
                            "pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body");
                    }
                    catch
                    {
                        continue;
                    }
                }
            }*/
        }

        public static bool createXml(string pathToFile)
        {
             

            return true;
        }

    }

    
}
