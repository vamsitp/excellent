namespace Excellent
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;

    using CommandLine;

    using Serilog;

    using SmartFormat;
    using SmartFormat.Core.Settings;

    internal class Program
    {
        private readonly static string TransformFormat = ConfigurationManager.AppSettings[nameof(TransformFormat)];

        private static void Main(string[] args)
        {
            SetLogger();
            SetSmartFormatting();
            var result = Parser.Default.ParseArguments<TransformOptions, MergeOptions, DiffOptions>(args)
                .MapResult(
                (TransformOptions opts) => Utils.Transform(opts.Input, !string.IsNullOrWhiteSpace(opts.Output) ? opts.Output : (Path.GetFileNameWithoutExtension(opts.Input) + ".txt"), TransformFormat),
                (MergeOptions opts) => Utils.Merge(opts.Inputs, !string.IsNullOrWhiteSpace(opts.Output) ? opts.Output : (string.Join("_", opts.Inputs.Select(Path.GetFileNameWithoutExtension)) + "_Merged.xlsx"), opts.KeepRight, opts.KeepLeft),
                (DiffOptions opts) => Utils.Diff(opts.Inputs, opts.Output), errs => HandleParseErrors(errs?.ToList()));
            if (result == 0)
            {
                Log.Information("Done!");
            }

            Log.CloseAndFlush();
            if (Debugger.IsAttached)
            {
                Console.ReadLine();
            }
        }

        private static int HandleParseErrors(List<Error> errs)
        {
            if (errs.Count > 1)
            {
                errs.ToList().ForEach(e => Log.Error(e.ToString()));
            }

            return errs.Count;
        }

        private static void SetLogger()
        {
            Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.Console(outputTemplate: "[{Level:u3}] {Message}{NewLine}").WriteTo.RollingFile("Excellent_{Date}.log", outputTemplate: "{Timestamp:dd-MMM-yyyy HH:mm:ss} | [{Level}] {Message}{NewLine}{Exception}").Enrich.FromLogContext().CreateLogger();
        }

        private static void SetSmartFormatting()
        {
            Smart.Default.Settings.ConvertCharacterStringLiterals = false;
            Smart.Default.Settings.CaseSensitivity = CaseSensitivityType.CaseInsensitive;
            Smart.Default.Settings.FormatErrorAction = ErrorAction.Ignore;
        }
    }
}
