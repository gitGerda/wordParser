using System;
using RaskroiParser.WorkComponents;
using System.IO;


namespace RaskroiParser
{
    class Program
    {
        static string pathToWorkDirectory;
        static string pathToParsedFiles;
        static void Main(string[] args)
        {
            //bool result = ArgumentsCheck.argumentsCheckFunc(args);

            if(true)
            {
                HideOrShow.hideOrShowFunc(args);
                /*pathToWorkDirectory = args[0];
                pathToParsedFiles = args[1];*/

                //Parsing.parsingFunc("2","2");

                pathToWorkDirectory = @"a:\WORK\PROJECTS\WordParsingScript\TestXML\";
                pathToParsedFiles = @"a:\WORK\PROJECTS\WordParsingScript\TestXML\parsed\";

                string pathToScript = Directory.GetCurrentDirectory();
                pathToScript = pathToScript + @"\..\..\..\..\ConvertDocToXML\bin\Release\ConvertDocToXML.exe";
                convertToXml.convert(pathToScript, pathToWorkDirectory, pathToParsedFiles);
            }

            

            Console.ReadLine();
        }

      
    }
}
