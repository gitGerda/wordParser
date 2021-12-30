using System;
using System.Collections.Generic;
using System.Text;

namespace RaskroiParser.WorkComponents
{
    public static class ArgumentsCheck
    {
        public static bool argumentsCheckFunc(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    throw new Exception("Error: Required arguments are not specified");
                }

                if (!System.IO.Directory.Exists(args[0]))
                {
                    throw new Exception("Error: No access or incorrect path to [PathToDocumentsDirectory]");
                }
                if (!System.IO.Directory.Exists(args[1]))
                {
                    throw new Exception("Error: No access or incorrect path to [PathToParsedFilesDirectory]");
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();

                return false;
            }

        }



    }
}
