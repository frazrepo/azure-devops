
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Clients;
using Microsoft.VisualStudio.Services.ReleaseManagement.WebApi.Contracts;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.IO;

/*
* Input : tokens.txt
* One line per variable#value
*/

namespace DevopsConsoleApi
{
    class Program
    {
        const string collectionUri = "https://dev.azure.com/org";
        const string personalAccessToken = "********************";
        const string projectName = "projectname";

        const int releaseDefinitionId = 1;
        const string environmentName = "Stage 1";
        const string tokenFilename = @"tokens.txt";

        static void Main(string[] args)
        {

            VssCredentials creds = new VssBasicCredential(string.Empty, personalAccessToken);

            // Connect to Azure DevOps Services
            VssConnection connection = new VssConnection(new Uri(collectionUri), creds);

            // Get a Release EndPoint
            ReleaseHttpClient releaseHttpClient = connection.GetClient<ReleaseHttpClient>();


            // Get data about a specific repository
            ReleaseDefinition definition = releaseHttpClient.GetReleaseDefinitionAsync(projectName, releaseDefinitionId).Result;


            //Determine if it is a release or a stage variable
            int environmentIndex = GetEnvironmentIndex(environmentName, definition.Environments);

            // Read Tokens from file (non secret variables)
            // Format per line : key#value
            var tokens = ReadTokensFromFile(tokenFilename);

            if (environmentIndex == -1)
            {
                foreach (var t in tokens)
                {
                    if (!definition.Variables.ContainsKey(t.Name))
                        definition.Variables.Add(t.Name, new ConfigurationVariableValue { Value = t.Value, IsSecret = false });
                    else
                        definition.Variables[t.Name] = new ConfigurationVariableValue { Value = t.Value, IsSecret = false };
                }
            }
            else
            {
                foreach (var t in tokens)
                {
                    if (!definition.Environments[environmentIndex].Variables.ContainsKey(t.Name))
                        definition.Environments[environmentIndex].Variables.Add(t.Name, new ConfigurationVariableValue { Value = t.Value, IsSecret = false });
                    else
                        definition.Environments[environmentIndex].Variables[t.Name] = new ConfigurationVariableValue { Value = t.Value, IsSecret = false };
                }
            }


            // Update Release Definition
            ReleaseDefinition updatedDefinition = releaseHttpClient.UpdateReleaseDefinitionAsync(definition, projectName).Result;

        }

        private static int GetEnvironmentIndex(string environmentName, IList<ReleaseDefinitionEnvironment> environments)
        {
            int result = -1;
            int counter = 0;
            foreach(var environment in environments)
            {
                if (environment.Name == environmentName)
                {
                    result = counter;
                    break;
                }

                counter++;
            }
            return result;
        }
        private static List<TokenValueHolder> ReadTokensFromFile(string fileName)
        {
            int counter = 0;
            string line;
            List<TokenValueHolder> results = new List<TokenValueHolder>();
 
            // Read the file and display it line by line.  
            StreamReader file =
                new StreamReader(fileName);

            char[] separator = { '#' };

            while ((line = file.ReadLine()) != null)
            {
                //Parse
                var tokensValues = line.Split(separator);
                var tokenValueHolder = new TokenValueHolder { Name = tokensValues[0], Value = tokensValues[1] };
                results.Add(tokenValueHolder);

                counter++;
            }

            return results;
        }

    }
}
