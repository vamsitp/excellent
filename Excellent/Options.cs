namespace Excellent
{
    using System.Collections.Generic;

    using CommandLine;
    using CommandLine.Text;

    public class Options
    {
        [Option('o', "output", Required = false, HelpText = "Output file.")]
        public string Output { get; set; }

        [Option('v', "verbose", Required = false, HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }
    }

    [Verb("transform", HelpText = "Excel file to be transformed based on the format specified in app config")]
    class TransformOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input file to be transformed.")]
        public string Input { get; set; }

        [Option('f', "format", Required = true, HelpText = "Transformation format (using Smart-Format).")]
        public string Format { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\n\nTRANSFORMATION - DEFAULT OUTPUT", new TransformOptions { Input = "Localizations.xlsx" });
                yield return new Example("\n\nTRANSFORMATION - SPECIFIC OUTPUT", new TransformOptions { Input = "Localizations.xlsx", Output = "Localizations.sql" });
                yield return new Example("\n\nTRANSFORMATION - SPECIFIC FORMAT", new TransformOptions { Input = "Localizations.xlsx", Output = "Localizations.sql", Format = "EXEC [dbo].[usp_InsertLocalizationData] @ResourceId = '{ResourceId}', @English = '{English}'" });
            }
        }
    }

    [Verb("merge", HelpText = "Excel files to be merged")]
    class MergeOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input files to be merged.")]
        public IEnumerable<string> Inputs { get; set; }

        [Option('l', "keep-left", Required = false, HelpText = "Retain the values from Left file when a duplicate exists.")]
        public bool KeepLeft { get; set; }

        [Option('r', "keep-right", Required = false, HelpText = "Retain the values from Right file when a duplicate exists.")]
        public bool KeepRight { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\nMERGE FILES", UnParserSettings.WithGroupSwitchesOnly(), new MergeOptions { Inputs = new[] { "Localizations_1.xlsx", "Localizations_2.xlsx" } });
            }
        }
    }

    [Verb("diff", HelpText = "Excel files to be diff'd")]
    class DiffOptions : Options
    {
        [Option('i', "input", Required = true, HelpText = "Input files to be diff'd.")]
        public IEnumerable<string> Inputs { get; set; }

        [Option('c', "sqlconn", Required = true, HelpText = "Sql Server Connection-string).")]
        public string SqlConnection { get; set; }

        [Usage(ApplicationAlias = "excellent.exe")]

        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("\nDIFF FILES", new[] { UnParserSettings.WithGroupSwitchesOnly(), UnParserSettings.WithUseEqualTokenOnly() }, new DiffOptions { Inputs = new[] { "Localizations_1.xlsx", "Localizations_2.xlsx" } });
                yield return new Example("\nDIFF FILES", new[] { UnParserSettings.WithGroupSwitchesOnly(), UnParserSettings.WithUseEqualTokenOnly() }, new DiffOptions { Inputs = new[] { "Localizations_1.xlsx", "Localization_Table.sql" }, SqlConnection = "Server=tcp:[serverName].database.windows.net;Database=myDataBase;User ID=[LoginForDb]@[serverName];Password=myPassword;Trusted_Connection=False;Encrypt=True;" });
            }
        }
    }
}
