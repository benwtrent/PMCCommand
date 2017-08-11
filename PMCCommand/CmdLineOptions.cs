// Copyright (c) 2017 Benjamin Trent. All rights reserved. See LICENSE file in project root

namespace PMCCommand
{
    using CommandLine;
    using CommandLine.Text;

    public class CmdLineOptions
    {
        [Option('n', "nugetcommand", Required = true, HelpText = "The NuGet package management console command to execute.")]
        public string NuGetCommand { get; set; }

        [Option('p', "project", Required = true, HelpText = "The full path of the csproj in which to run the command.")]
        public string ProjectPath { get; set; }

        [Option('v', "vsversion", Required = false, HelpText = "The VisualStudio version for DTE interaction", DefaultValue = "14.0")]
        public string VisualStudioVersion { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("PMCCommand", "1.0.0"),
                Copyright = new CopyrightInfo("Benjamin Trent", 2017),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };
            help.AddOptions(this);
            help.AddPostOptionsLine("Example: PMCCommand --nugetcommand \"Update-Package Newtonsoft.Json\" --project \"C:\\Foo\\Bar\\foobar.csproj\"");
            return help;
        }
    }
}
