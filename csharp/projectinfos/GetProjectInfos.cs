
using CommandLine;
using CommandLine.Text;
using GetProjectInfos;
using Microsoft.TeamFoundation.Build.WebApi;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.Graph.Client;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Clients;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Contracts;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;

namespace UpdateTokens
{
    class Program
    {
        const string collectionUri = "https://dev.azure.com/organization";
        const string personalAccessToken = "******";

        public class Options
        {
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }

            //[Option("pat", Required = false, HelpText = "Personal Access Token")]
            //public string Pat { get; set; }

            [Usage(ApplicationAlias = "UpdateTokens")]
            public static IEnumerable<Example> Examples
            {
                get
                {
                    return new List<Example>() {
                        new Example("Common Usage", new Options { Verbose = true})
                     };
                }
            }
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

            //GraphClient
            GraphHttpClient graphClient = new GraphHttpClient(new Uri(collectionUri), creds);

            // Get Project
            ProjectHttpClient projectClient = connection.GetClient<ProjectHttpClient>();
            var projects = projectClient.GetProjects().Result;

            ReleaseHttpClient releaseHttpClient = connection.GetClient<ReleaseHttpClient>();
            BuildHttpClient buildClient = connection.GetClient<BuildHttpClient>();

            // Results
            List<ProjectExtractInfo> projectExtracts = new List<ProjectExtractInfo>();
            List<BuildExtractInfo> buildExtracts= new List<BuildExtractInfo>();
            List<ReleaseExtractInfo> releaseExtracts = new List<ReleaseExtractInfo>();

            foreach (var p in projects)
            {
                Console.WriteLine($"Processing {p.Name} ...");

                //Builds
                var builds = buildClient.GetFullDefinitionsAsync(p.Id).Result;
                foreach (var b in builds)
                {
                    buildExtracts.Add(new BuildExtractInfo
                    {
                        Name = b.Name,
                        ProjectName = p.Name,
                        RepoName = b.Repository.Name,
                        RepoType = b.Repository.Type,
                        DefaultAgent = b.Queue.Pool.Name,
                        TriggerType = b.Triggers.Count >0 ?  b.Triggers[0].TriggerType.ToString() : "",
                        TriggerBranchFilters = b.Triggers.Count >0 ? string.Join(",",((ContinuousIntegrationTrigger)b.Triggers.First()).BranchFilters.ToArray()) : "",
                        Url = b.Url,
                        Steps = b.Process!=null ? ((DesignerProcess)b.Process).Phases[0].Steps.Select(s => s.DisplayName).ToList() : new List<string>()
                    }) ;
                }

                //Releases
                var releases = releaseHttpClient.GetReleaseDefinitionsAsync(p.Name, expand: ReleaseDefinitionExpands.Environments).Result;
                foreach (var r in releases)
                {
                    releaseExtracts.Add(new ReleaseExtractInfo
                    {
                        Name = r.Name,
                        ProjectName = p.Name
                    });
                }

                projectExtracts.Add(new ProjectExtractInfo { Name = p.Name, BuildCount = builds.Count, ReleaseCount = releases.Count });
            }


            //Save Results
            ExcelSerializer.Serialize(projectExtracts, buildExtracts, releaseExtracts);

            Console.ReadLine();

        }
    }
}
