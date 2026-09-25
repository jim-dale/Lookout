namespace Lookout.CommandBuilders;

using System.CommandLine;
using Lookout.Commands;
using Lookout.Options;
using OutlookServices.Options;

/// <summary>
/// Builds the command used to export attachments.
/// </summary>
internal static class ExportAttachmentsCommandBuilder
{
    /// <summary>
    /// Builds the ExportAttachments command and configures its hosted service.
    /// </summary>
    /// <param name="hostBuilder">The host application builder used to configure the command.</param>
    /// <returns>The configured command.</returns>
    internal static Command Build(HostApplicationBuilder hostBuilder)
    {
        Option<string> storeDescriptorOption = CommonOptions.GetStoreDescriptorOption();
        Option<FileInfo[]> datafilesOption = CommonOptions.GetDatafilesOption();

        Option<DirectoryInfo> outputFolderOption = new("--output", "-o")
        {
            Description = "Path of folder to save attachments to.",
            Arity = ArgumentArity.ZeroOrOne,
            DefaultValueFactory = _ => new DirectoryInfo(@".\attachments"),
        };

        Option<int> maxRecurseOption = new("--max-recursion", "-m")
        {
            Description = "Maximum folder recursion level.",
            DefaultValueFactory = _ => ExportAttachmentsOptions.DefaultMaxRecurseLevels,
        };

        Option<bool> whatIfOption = new("--what-if", "-w")
        {
            Description = "Shows the effact of the command.",
        };

        Command result = new("attachments")
        {
            Options = { storeDescriptorOption, datafilesOption, outputFolderOption, maxRecurseOption, whatIfOption, },
        };

        result.SetAction(async parseResult =>
        {
            // process any command line overrides of configured options
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

            // process any command line overrides of configured options
            hostBuilder.Services.PostConfigure<ExportAttachmentsOptions>(opts =>
            {
                DirectoryInfo? outputFolder = parseResult.GetValue(outputFolderOption);
                if (outputFolder is not null)
                {
                    opts.TargetFolder = outputFolder;
                }

                int? maxRecurseLevel = parseResult.GetValue(maxRecurseOption);
                if (maxRecurseLevel.HasValue)
                {
                    opts.MaxRecurseLevel = maxRecurseLevel.Value;
                }

                bool? whatIf = parseResult.GetValue(whatIfOption);
                if (whatIf.HasValue)
                {
                    opts.WhatIf = whatIf.Value;
                }
            });

            hostBuilder.Services.AddHostedService<ExportAttachmentsCommand>();

            using IHost host = hostBuilder.Build();

            await host.RunAsync();
            return 0;
        });

        return result;
    }
}
