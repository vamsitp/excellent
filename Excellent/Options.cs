namespace Excellent
{
    using System.Collections.Generic;

    using CommandLine;
    using CommandLine.Text;

    class Options
    {
        [Option('o', "output", Required = false, HelpText = "Output file.")]
        public string OutputFile { get; set; }

        [Option('v', "verbose", Required = false, HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }
    }

    [Verb("transform", HelpText = "Excel file to be transformed based on the format specified in app config")]
    class TransformOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input file to be transformed.")]
        public string InputFile { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\n\nTRANSFORMATION - DEFAULT OUTPUT", new TransformOptions { InputFile = "Localizations.xlsx" });
                yield return new Example("\n\nTRANSFORMATION - SPECIFIC OUTPUT", new TransformOptions { InputFile = "Localizations.xlsx", OutputFile = "Localizations.sql" });
            }
        }
    }

    [Verb("merge", HelpText = "Excel files to be merged")]
    class MergeOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input files to be merged.")]
        public IEnumerable<string> MergeFiles { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\nMERGE FILES", UnParserSettings.WithGroupSwitchesOnly(), new MergeOptions { MergeFiles = new[] { "File1.xlsx", "File2.xlsx" } });
            }
        }
    }

    [Verb("diff", HelpText = "Excel files to be diff'd")]
    class DiffOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input files to be diff'd.")]
        public IEnumerable<string> DiffFiles { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\nDIFF FILES", new[] { UnParserSettings.WithGroupSwitchesOnly(), UnParserSettings.WithUseEqualTokenOnly() }, new DiffOptions { DiffFiles = new[] { "File1.xlsx", "File2.xlsx" } });
            }
        }
    }
}
