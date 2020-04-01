
using CommandLine;
using CommandLine.Text;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Clients;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Contracts;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.IO;

namespace GetProjectInfos
{
    class Program
    {
        const string collectionUri = "https://dev.azure.com/Company";
        const string personalAccessToken = "pat";

        public class Options
        {
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }

            [Option('p', "project", Required = true, HelpText = "Set project name")]
            public string ProjectName { get; set; }

            [Option('r', "releaseid", Required = true, HelpText = "Set release id")]
            public int ReleaseId { get; set; }
        }

        static void Main(string[] args)
        {
            try
            {
                CommandLine.Parser.Default.ParseArguments<Options>(args)
                    .WithParsed(RunOptions)
                    .WithNotParsed(HandleParseError);

            }
            catch (Exception e)
            {
                Console.WriteLine("Error", e.Message);
            }
        }

        static void HandleParseError(IEnumerable<Error> errs)
        {
            //handle errors
        }

        static void RunOptions(Options opts)
        {

            VssCredentials creds = new VssBasicCredential(string.Empty, personalAccessToken);

            // Connect to Azure DevOps Services
            VssConnection connection = new VssConnection(new Uri(collectionUri), creds);

            // Get a Release EndPoint
            ReleaseHttpClient releaseHttpClient = connection.GetClient<ReleaseHttpClient>();

            // Get data about a specific repository
            ReleaseDefinition definition = releaseHttpClient.GetReleaseDefinitionAsync(opts.ProjectName, opts.ReleaseId).Result;

            List<ReleaseDefinitionEnvironment> environments = definition.Environments as List<ReleaseDefinitionEnvironment>;
            environments.Sort();


            foreach (var env in environments)
            {
                Console.WriteLine();
                Console.WriteLine($"## {env.Name}");
                Console.WriteLine();

                foreach (var variable in env.Variables)
                {
                    Console.WriteLine($"{variable.Key}#{variable.Value.Value}");
                }

            }
        }
    }
}
