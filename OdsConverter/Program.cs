using System;
using System.IO;

namespace OdsConverter
{
    class Program
    {
        static int Main(string[] args)
        {
            //Parse args
            if (args.Length < 2)
                return PrintUsage();
            string scalcLocation = args[0];
            if (!File.Exists(scalcLocation))
                return PrintUsage();
            string inputDirectory = args[1];
            if (!Directory.Exists(inputDirectory))
                return PrintUsage();
            string outputDirectory = inputDirectory;
            if (args.Length > 2)
                outputDirectory = args[2];

            Converter converter = new Converter(scalcLocation, inputDirectory, outputDirectory, "*.ods", "xls", Log);
            converter.Convert();
            return 0;
        }

        static void Log(string msg)
        {
            Console.WriteLine(msg);
        }

        private static int PrintUsage()
        {
            Console.WriteLine("Usage: OdsConverter [path to scalc.exe] [input dir] ([output dir])");
            return 1;
        }
    }
}
