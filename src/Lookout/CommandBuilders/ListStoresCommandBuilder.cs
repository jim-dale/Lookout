namespace Lookout.CommandBuilders;

using System.CommandLine;
using Lookout.Commands;

/// <summary>
/// Creates the 'stores' CLI command that registers ListStoresCommand as a hosted service and runs the host asynchronously.
/// </summary>
internal class ListStoresCommandBuilder
{
    /// <summary>
    /// Creates a 'stores' command that, when invoked, registers the ListStoresCommand hosted service, builds the host,
    /// and runs it asynchronously.
    /// </summary>
    /// <param name="hostBuilder">HostApplicationBuilder used to register services and construct the IHost for command execution.</param>
    /// <returns>A configured Command that will register and run the host to list stores and return an exit code.</returns>
    public static Command Build(HostApplicationBuilder hostBuilder)
    {
        Command result = new("stores");
        result.SetAction(async parseResult =>
        {
            hostBuilder.Services.AddHostedService<ListStoresCommand>();

            using IHost host = hostBuilder.Build();

            await host.RunAsync();
            return 0;
        });

        return result;
    }
}
