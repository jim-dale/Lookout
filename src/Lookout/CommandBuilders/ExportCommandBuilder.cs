namespace Lookout.CommandBuilders;

using System.CommandLine;

/// <summary>
/// Builds the commands used to export entities from Outlook Stores.
/// </summary>
internal static class ExportCommandBuilder
{
    /// <summary>
    /// Creates the 'export' root command that groups Outlook export subcommands.
    /// </summary>
    /// <param name="hostBuilder">Host application builder used to construct and configure subcommands and resolve services.</param>
    /// <returns>A Command named 'export' configured with Outlook export subcommands.</returns>
    public static Command Build(HostApplicationBuilder hostBuilder)
    {
        Command exportAttachments = ExportAttachmentsCommandBuilder.Build(hostBuilder);

        Command result = new("export", "Export Outlook entities")
        {
            Subcommands = { exportAttachments },
        };

        return result;
    }
}
