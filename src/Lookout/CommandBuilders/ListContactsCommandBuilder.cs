namespace Lookout.CommandBuilders;

using System.CommandLine;
using Lookout.Commands;
using Lookout.Options;
using OutlookServices.Options;

/// <summary>
/// Builds the command used to list Outlook contacts.
/// </summary>
internal class ListContactsCommandBuilder
{
    /// <summary>
    /// Builds the ListContacts command and configures its hosted service.
    /// </summary>
    /// <param name="hostBuilder">The host application builder used to configure the command.</param>
    /// <returns>The configured command.</returns>
    public static Command Build(HostApplicationBuilder hostBuilder)
    {
        Option<string> storeDescriptorOption = CommonOptions.GetStoreDescriptorOption();
        Option<FileInfo[]> datafilesOption = CommonOptions.GetDatafilesOption();

        Option<string[]> propertyNamesOption = new("--property", "-p")
        {
            Description = "Outlook Contact item property name.",
            Arity = ArgumentArity.OneOrMore,
        };

        Option<int> maxRecurseOption = new("--max-recursion", "-m")
        {
            Description = "Maximum folder recursion level.",
            DefaultValueFactory = _ => ListContactsOptions.DefaultMaxRecurseLevels,
        };

        Command result = new("contacts")
        {
            Options = { storeDescriptorOption, datafilesOption, propertyNamesOption, maxRecurseOption, },
        };

        result.SetAction(async parseResult =>
        {
            hostBuilder.Services.PostConfigure<OutlookServiceOptions>(opts =>
            {
                string? storeDescriptor = parseResult.GetValue(storeDescriptorOption);
                if (!string.IsNullOrWhiteSpace(storeDescriptor))
                {
                    opts.StoreDescriptor = storeDescriptor;
                }

                FileInfo[]? dataFiles = parseResult.GetValue(datafilesOption);
                if (dataFiles is not null)
                {
                    opts.AdditionalDataFiles.AddRange(dataFiles);
                }
            });

            hostBuilder.Services.PostConfigure<ListContactsOptions>(opts =>
            {
                int? maxRecurseLevel = parseResult.GetValue(maxRecurseOption);
                if (maxRecurseLevel.HasValue)
                {
                    opts.MaxRecurseLevel = maxRecurseLevel.Value;
                }

                string[]? propertyNames = parseResult.GetValue(propertyNamesOption);
                if (propertyNames is not null)
                {
                    opts.PropertyNames = propertyNames;
                }
            });

            hostBuilder.Services.AddHostedService<ListContactsCommand>();

            using IHost host = hostBuilder.Build();

            await host.RunAsync();
            return 0;
        });

        return result;
    }
}
