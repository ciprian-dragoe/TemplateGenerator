using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateGenerator
{
    public class MainClass
    {
        // args[0] -> pathToDBF
        // args[1] -> pathToDocx
        // args[2] -> pathToSaveDir
        public static void Main(string[] args)
        {
            if (args.Length<3)
            {
                throw new Exception("E nevoie de minim 3 argumente pentru a rula programul.");
            }
            TemplateManager.generateDocxTemplates(args[0], args[1], args[2]);
        }
    }
}
