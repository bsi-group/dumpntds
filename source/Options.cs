using CommandLine;

namespace dumpntds
{
    /// <summary>
    /// Internal class used for the command line parsing
    /// </summary>
    internal class Options
    {
        [Option('n', "ntds", Required = true, Default = "", HelpText = "Path to ntds.dit file")]
        public string Ntds { get; set; }

        [Option('t', "type", Required = false, Default = ExportType.Csv, HelpText = "Export type")]
        public ExportType ExportType { get; set; }
    }
}
