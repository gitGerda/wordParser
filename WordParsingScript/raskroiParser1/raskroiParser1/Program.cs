using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using raskroiParser1.components;

namespace raskroiParser1
{
    class Program
    {
        static void Main(string[] args)
        {
            if(argsCheck.argumentsCheckFunc(args))
            {
                int h = hideOrShow.hideOrShowFunc(args);
                parser.parserFunc(args[0],args[1]);   
            }
        }
    }
}
