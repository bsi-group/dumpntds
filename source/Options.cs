using System;
using CommandLine;
using CommandLine.Text;

namespace dumpntds
{
    /// <summary>
    /// Internal class used for the command line parsing
    /// </summary>
    internal class Options
    {
        [ParserState]
        public IParserState LastParserState { get; set; }

        [Option('n', "ntds", Required = true, DefaultValue = "", HelpText = "Path to ntds.dit file")]
        public string Ntds { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Copyright = new CopyrightInfo("Info-Assure", 2015),
                AdditionalNewLineAfterOption = false,
                AddDashesToOption = true
            };

            help.AddPreOptionsLine("Usage: dumpntds -m p -n ntds.dit");
            help.AddOptions(this);

            if( this.LastParserState != null)
            {
                var errors = help.RenderParsingErrorsText(this, 0); // indent with two spaces
                if (!string.IsNullOrEmpty(errors))
                {
                    help.AddPreOptionsLine(string.Concat(Environment.NewLine, "ERROR(S):"));
                    help.AddPreOptionsLine(errors);
                }
            }

            return help;
        }
    }
}
