namespace Lookout.CommandBuilders;

using System.CommandLine;

/// <summary>
/// Provides options shared by multiple commands.
/// </summary>
internal static class CommonOptions
{
    /// <summary>
    /// Creates the option used to select stores for an operation.
    /// </summary>
    /// <returns>An option that specifies the stores to select.</returns>
    internal static Option<string> GetStoreDescriptorOption()
    {
        return new("--store", "-s")
        {
            Description = "Full or partial store name, id, or path. Can also be '*' to match all stores, or 'default' to match the default store.",
            Arity = ArgumentArity.ZeroOrOne,
            DefaultValueFactory = _ => "*",
        };
    }

    /// <summary>
    /// Creates the option used to specify additional PST data files to load.
    /// </summary>
    /// <returns>An option that specifies additional PST data files to load.</returns>
    internal static Option<FileInfo[]> GetDatafilesOption()
    {
        return new("--data-file", "-d")
        {
            Description = "Additional (pst) data files to load.",
            Arity = ArgumentArity.ZeroOrMore,
        };
    }
}
