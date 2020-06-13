using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace FunctionGraphAPI
{
  public static class GetGraphData
  {
    [FunctionName("GetGraphData")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
        ILogger log)
    {
      log.LogInformation("C# HTTP trigger function processed a request.");

      string clientId = Environment.GetEnvironmentVariable("ClientId");
      string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");
      string tenantId = Environment.GetEnvironmentVariable("TenantId");
      string authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

      var cca = ConfidentialClientApplicationBuilder.Create(clientId)
        .WithClientSecret(clientSecret)
        .WithAuthority(authority)
        .Build();

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      MSALAuthenticationProvider authenticationProvider = new MSALAuthenticationProvider(cca, scopes.ToArray());
      GraphServiceClient graphServiceClient = new GraphServiceClient(authenticationProvider);

      List<QueryOption> options = new List<QueryOption>
      {
        new QueryOption("$top", "2")
      };

      var graphResult = await graphServiceClient.Users.Request(options).GetAsync();//.Result;

      return new OkObjectResult(graphResult);
    }
  }
}
