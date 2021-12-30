using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Text.Json;


namespace raskroiParser1.components
{
    public static class parser
    {
        public static void parserFunc(string pathToWorkDir, string pathToParsed)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;

            DirectoryInfo dirInfo = new DirectoryInfo(pathToWorkDir);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");

            word.Visible = false;
            word.ScreenUpdating = false;

            XmlDocument xmlDoc = new XmlDocument();

            foreach (FileInfo wordFile in wordFiles)
            {
                string pathToFileWord = wordFile.DirectoryName+"\\";
                string FullPathWord = wordFile.FullName;

                string fileName = FullPathWord.Replace(pathToFileWord, "");
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

                        string jsonDoc = @"{ ""Поднестинг"":";

                        string podnestingCell = doc.Tables[1].Rows[1].Cells[2].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref podnestingCell);
                        jsonDoc += @"""" + podnestingCell + @""",";

                        string dateCell = doc.Tables[1].Rows[1].Cells[3].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref dateCell);
                        jsonDoc += @"""Дата"":""" + dateCell + @""",";

                        string materialCell = doc.Tables[1].Rows[3].Cells[2].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref materialCell);
                        jsonDoc += @"""Материал"":""" + materialCell + @""",";

                        Microsoft.Office.Interop.Word.Table tbl2 = doc.Tables[2].Cell(1, 1).Tables[1];

                        string listCountCell = tbl2.Rows[2].Cells[2].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref listCountCell);
                        jsonDoc += @"""КоличествоЛистов"":""" + listCountCell + @""",";

                        string detCountCell = tbl2.Rows[2].Cells[3].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref detCountCell);
                        jsonDoc += @"""КоличествоДеталей"":""" + detCountCell + @""",";

                        string effecCell = tbl2.Rows[2].Cells[4].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref effecCell);
                        jsonDoc += @"""Эффективность"":""" + effecCell + @""",";

                        string othCell = tbl2.Rows[2].Cells[5].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref othCell);
                        jsonDoc += @"""Отход"":""" + othCell + @""",";

                        string timeCell = tbl2.Rows[2].Cells[6].Range.Text;
                        delEscapeChars.delEscapeCharsFunc(ref timeCell);
                        jsonDoc += @"""РасчетноеВремя"":""" + timeCell + @""",";

                        jsonDoc += @"""Детали"":[";

                        bool headerFlag = true;
                        foreach (Microsoft.Office.Interop.Word.Row row in doc.Tables[3].Rows)
                        {
                            if (!headerFlag)
                            {
                                string detfileCell = row.Cells[2].Range.Text;
                                delEscapeChars.delEscapeCharsFunc(ref detfileCell);
                                jsonDoc += @"{""Файл"":""" + detfileCell + @""",";

                                string detCountLocalCell = row.Cells[6].Range.Text;
                                delEscapeChars.delEscapeCharsFunc(ref detCountLocalCell);
                                jsonDoc += @"""Количество"":""" + detCountLocalCell + @""",";

                                string orderedCell = row.Cells[7].Range.Text;
                                delEscapeChars.delEscapeCharsFunc(ref orderedCell);
                                jsonDoc += @"""Заказанное"":""" + orderedCell + @"""}";

                                if (row.Next != null)
                                {
                                    jsonDoc += ",";
                                }
                                else
                                {
                                    jsonDoc += "]}";
                                }
                            }
                            headerFlag = false;
                        }

                        /*string saveStrPath = pathToFileWord + "[parsed]" + fileName;*/

                        //doc.SaveAs(saveStrPath);

                        ((Microsoft.Office.Interop.Word._Document)doc).Close();
                        doc = null;
                        File.Delete(pathToFileWord + fileName);


                        /*Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("_success: " + saveStrPath);*/
                        


                        string target = "";
                        string pattern = ".doc";
                        System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(pattern);
                        fileName = reg.Replace(fileName, target);

                        string jsonfileName = pathToParsed + fileName + ".json";
                        using (FileStream fs = new FileStream(jsonfileName, FileMode.Create))
                        {
                            byte[] array = System.Text.Encoding.UTF8.GetBytes(jsonDoc);
                            fs.Write(array,0,array.Length);
                        }

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("_success: " + jsonfileName);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);

                        /*string errorDesc = "\n#" + DateTime.Now.ToString() + "--->" + ex.Message;
                        File.AppendAllText(Environment.SystemDirectory + @"\logs.txt", errorDesc);*/

                        continue;
                    }
                }
            }
            word.Quit();

            Console.ResetColor();
        }
    }
}
