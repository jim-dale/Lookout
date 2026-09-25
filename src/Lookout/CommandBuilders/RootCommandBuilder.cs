namespace Lookout.CommandBuilders;

using System.CommandLine;

/// <summary>
/// Provides a builder for the application's RootCommand and registers the List and Export subcommands.
/// </summary>
/// <remarks>Constructs subcommands using the provided HostApplicationBuilder via ListCommandBuilder and
/// ExportCommandBuilder, returning a configured RootCommand.</remarks>
internal static class RootCommandBuilder
{
    /// <summary>
    /// Creates a RootCommand titled "Outlook Services" and registers the subcommands.
    /// </summary>
    /// <param name="hostBuilder">HostApplicationBuilder used to construct and configure the subordinate command builders.</param>
    /// <returns>A RootCommand configured with the list and export subcommands.</returns>
    public static RootCommand Build(HostApplicationBuilder hostBuilder)
    {
        Command listCommand = ListCommandBuilder.Build(hostBuilder);
        Command exportCommand = ExportCommandBuilder.Build(hostBuilder);

        RootCommand result = new("Outlook Services")
        {
            Subcommands = { listCommand, exportCommand },
        };

        return result;
    }
}
