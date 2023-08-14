using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Orchestration;
using Microsoft.SemanticKernel.SkillDefinition;
using Microsoft.SemanticKernel.Skills.MsGraph;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Client;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.CredentialManagers;
using DayOfWeek = System.DayOfWeek;

namespace SemanticKernelConsoleApp;
public sealed class Program{
     public static async Task Main()
    {
        // Load configuration
        IConfigurationRoot configuration = new ConfigurationBuilder()
            .AddJsonFile(path: "appsettings.json", optional: false, reloadOnChange: true)
            .AddJsonFile(path: "appsettings.Development.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .AddUserSecrets<Program>()
            .Build();

        // Get the MsGraph configuration from appsettings
        MsGraphConfiguration? msGraphConfiguration = configuration.GetRequiredSection("MsGraph").Get<MsGraphConfiguration>();
        configuration.GetSection("MsGraph:Scopes").Bind(msGraphConfiguration.Scopes);
        var defaultCompletionServiceId = configuration["DefaultCompletionServiceId"];

        // Add authentication handler.
        IList<DelegatingHandler> handlers = GraphClientFactory.CreateDefaultHandlers(
            CreateAuthenticationProvider(await LocalUserMSALCredentialManager.CreateAsync(), msGraphConfiguration));

        // Create the Graph client.
        using HttpClient httpClient = GraphClientFactory.Create(handlers);
        GraphServiceClient graphServiceClient = new(httpClient);

        // Initialize SK Graph API Skills we'll be using in the plan.
        CloudDriveSkill oneDriveSkill = new(new OneDriveConnector(graphServiceClient));
        TaskListSkill todoSkill = new(new MicrosoftToDoConnector(graphServiceClient));
        EmailSkill outlookSkill = new(new OutlookMailConnector(graphServiceClient));
        CalendarSkill calendarSkill = new(new OutlookCalendarConnector(graphServiceClient));
        
        // Initialize the Semantic Kernel and and register connections with OpenAI/Azure OpenAI instances.
        KernelBuilder builder = Kernel.Builder;

       
        AzureOpenAIConfiguration? azureOpenAIConfiguration = configuration.GetSection("AzureOpenAI").Get<AzureOpenAIConfiguration>();
            
        builder.WithAzureChatCompletionService(
            deploymentName: azureOpenAIConfiguration.DeploymentName,
            endpoint: azureOpenAIConfiguration.Endpoint,
            apiKey: azureOpenAIConfiguration.ApiKey,
            serviceId: azureOpenAIConfiguration.ServiceId,
            setAsDefault: azureOpenAIConfiguration.ServiceId == defaultCompletionServiceId);
            
        // Add the skills to the kernel.
        IKernel sk = builder.Build();

        var onedrive = sk.ImportSkill(oneDriveSkill, "onedrive");
        var todo = sk.ImportSkill(todoSkill, "todo");
        var outlook = sk.ImportSkill(outlookSkill, "outlook");
        var calendar = sk.ImportSkill(calendarSkill, "calendar");

        // Import skills from the local directory.
        IDictionary<string, ISKFunction> summarizeSkills =
            sk.ImportSemanticSkillFromDirectory("./skills", "SummarizeSkill");

        // Get file path from appsettings
        var pathToFile = configuration["OneDrivePathToFile"];
        
        // Summarize the file
        SKContext fileContentResult = await sk.RunAsync(pathToFile, onedrive["GetFileContent"], summarizeSkills["Summarize"]);

        // Get the summary
        string fileSummary = fileContentResult.Result;

        // Get my email address
        SKContext emailAddressResult = await sk.RunAsync(string.Empty, outlook["GetMyEmailAddress"]);
        string myEmailAddress = emailAddressResult.Result;

        // Create a sharelink to the file
        SKContext fileLinkResult = await sk.RunAsync(pathToFile, onedrive["CreateLink"]);
        string fileLink = fileLinkResult.Result;

        // Send me an email with the summary and a link to the file
        ContextVariables emailMemory = new($"{fileSummary}{Environment.NewLine}{Environment.NewLine}{fileLink}");
        emailMemory.Set(EmailSkill.Parameters.Recipients, myEmailAddress);
        emailMemory.Set(EmailSkill.Parameters.Subject, $"Summary of {pathToFile}");
        await sk.RunAsync(emailMemory, outlook["SendEmail"]);

        Console.WriteLine($"Sent email to {myEmailAddress} with summary of {pathToFile}.");

        // Add a reminder on ToDo to follow-up next week
        ContextVariables followUpTaskMemory = new($"Follow-up about {pathToFile}.");
        DateTimeOffset nextMonday = TaskListSkill.GetNextDayOfWeek(DayOfWeek.Monday, TimeSpan.FromHours(9));
        followUpTaskMemory.Set(TaskListSkill.Parameters.Reminder, nextMonday.ToString("o"));
        await sk.RunAsync(followUpTaskMemory, todo["AddTask"]);

        Console.WriteLine($"Added a reminder on ToDo to follow-up next week about {pathToFile}.");

        // Add a calendar event to follow-up next week
        ContextVariables followUpEventMemory = new($"Follow-up about {pathToFile}.");
        followUpEventMemory.Set(CalendarSkill.Parameters.Start, nextMonday.ToString("o"));
        followUpEventMemory.Set(CalendarSkill.Parameters.End, nextMonday.AddHours(1).ToString("o"));
        await sk.RunAsync(followUpEventMemory, calendar["AddEvent"]);

        Console.WriteLine($"Added a calendar event to follow-up next week about {pathToFile}.");
        
    }

    /// Create a delegated authentication callback for the Graph API client.
    private static DelegateAuthenticationProvider CreateAuthenticationProvider(
        LocalUserMSALCredentialManager credentialManager,
        MsGraphConfiguration config)
        => new(
            async (requestMessage) =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue(
                    scheme: "bearer",
                    parameter: await credentialManager.GetTokenAsync(
                        config.ClientId,
                        config.TenantId,
                        config.Scopes.ToArray(),
                        config.RedirectUri));
            });
}