# Installation
# pip install azure-devops
# https://github.com/Microsoft/azure-devops-python-api
# You can specify the proxy by setting the environment variable HTTPS_PROXY or HTTP_PROXY to the url of your proxy.

from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
import pprint

# Fill in with your personal access token and org URL
personal_access_token = 'PAT'
organization_url = 'https://dev.azure.com/org'

# Create a connection to the org
credentials = BasicAuthentication('', personal_access_token)
connection = Connection(base_url=organization_url, creds=credentials)

# Get a client (the "core" client provides access to projects, teams, etc)
core_client = connection.clients.get_core_client()

# Get the first page of projects
get_projects_response = core_client.get_projects()

# Display all projects inside an organization
index=0
while get_projects_response is not None:
    for project in get_projects_response.value:
        pprint.pprint("[" + str(index) + "] " + project.name)
        index += 1
