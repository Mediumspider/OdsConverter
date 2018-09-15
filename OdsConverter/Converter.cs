using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OdsConverter
{
    internal class Converter
    {
        private string m_ScalcLocation;
        private string m_InputDirectory;
        private string m_OutputDirectory;
        private string m_WildCard;
        private string m_OutputType;
        private Action<string> m_Log;

        public Converter(string scalcLocation, string inputDirectory, string outputDirectory, string wildCard, string output, Action<string> log = null)
        {
            m_ScalcLocation = scalcLocation;
            m_InputDirectory = inputDirectory;
            m_OutputDirectory = outputDirectory;
            m_WildCard = wildCard;
            m_OutputType = output;
            m_Log = log ?? new Action<string>((msg) => Console.WriteLine(msg));
        }

        internal void Convert()
        {
            HandleDirectory(m_InputDirectory, m_OutputDirectory);
            
        }

        private void HandleDirectory(string inputDirectory, string outputDirectory)
        {
            string[] files = Directory.GetFiles(inputDirectory, m_WildCard);
            string[] directories = Directory.GetDirectories(inputDirectory);
            if (files.Any())
            {
                m_Log(string.Format("Converting files in {0} => {1}", inputDirectory, outputDirectory));
                Directory.CreateDirectory(outputDirectory);
            }
            foreach (string file in files)
            {
                m_Log(string.Format("Converting {0}", file));
                string arguments = string.Format("--convert-to \"{0}\" --outdir \"{1}\" \"{2}\"", m_OutputType, outputDirectory, file);
                ProcessStartInfo startInfo = new ProcessStartInfo(m_ScalcLocation, arguments);
                startInfo.WorkingDirectory = Path.GetDirectoryName(m_ScalcLocation);
                Process scalc = Process.Start(startInfo);
                scalc.WaitForExit();
            }
            foreach (string directory in directories)
            {
                string subDir = Path.Combine(outputDirectory, Path.GetFileName(directory));
                HandleDirectory(directory, subDir);
            }
        }
    }
}