using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Xml;


namespace ConvertDocToXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToWorkDir = @"a:\WORK\PROJECTS\WordParsingScript\TestXML\";
            string pathToParsed = @"a:\WORK\PROJECTS\WordParsingScript\TestXML\parsed\";

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;

            DirectoryInfo dirInfo = new DirectoryInfo(pathToWorkDir);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");

            word.Visible = false;
            word.ScreenUpdating = false;

            XmlDocument xmlDoc = new XmlDocument();

            foreach (FileInfo wordFile in wordFiles)
            {
                string fileName = wordFile.Name;
                if (!fileName.Contains("[parsed]"))
                {
                    try
                    {
                        Object filenamePath = (Object)wordFile.FullName;
                        Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filenamePath, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        doc.Activate();
                        //doc.Tables[1].Cell(1, 1).Range.Text;

                        object outputFileName = wordFile.FullName.Replace(".doc", ".xml");

                        object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML;
                        doc.SaveAs(ref outputFileName, ref fileFormat, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                        ((Microsoft.Office.Interop.Word._Document)doc).Close(ref saveChanges);
                        doc = null;



                        doc = word.Documents.Open(ref filenamePath, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        doc.Activate();
                        object outputFileName2 = wordFile.FullName.Insert(0, "[parsed]");
                        doc.SaveAs(ref outputFileName2, ref fileFormat, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                     ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        /*xmlDoc.Load(outputFileName.ToString());
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                        nsmgr.AddNamespace("w", "http://schemas.microsoft.com/office/word/2003/wordml");
                        nsmgr.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");*/

                        /*XmlNodeList documentBody = xmlDoc.SelectNodes("//pkg:package/" +
                            "pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body",nsmgr);*/

                        /*foreach(XmlNode n in documentBody)
                        {

                        }*/
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        continue;

                    }
                }
            }
            word.Quit();
        }
    }
}
