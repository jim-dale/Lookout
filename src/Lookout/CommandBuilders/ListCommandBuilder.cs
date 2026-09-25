namespace Lookout.CommandBuilders;

using System.CommandLine;

/// <summary>
/// Builds the commands used to list entities in Outlook Stores.
/// </summary>
internal static class ListCommandBuilder
{
    /// <summary>
    /// Creates the 'list' root command that groups Outlook list subcommands.
    /// </summary>
    /// <param name="hostBuilder">Host application builder used to construct and configure subcommands and resolve services.</param>
    /// <returns>A Command named 'list' configured with Outlook list subcommands.</returns>
    public static Command Build(HostApplicationBuilder hostBuilder)
    {
        Command listStores = ListStoresCommandBuilder.Build(hostBuilder);
        Command listContacts = ListContactsCommandBuilder.Build(hostBuilder);

        Command result = new("list", "List entities in Outlook")
        {
            Subcommands = { listStores, listContacts, },
        };

        return result;
    }
}
