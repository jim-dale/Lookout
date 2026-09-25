namespace Lookout;

using System;
using System.CommandLine;
using System.Text;
using System.Threading.Tasks;
using Lookout.CommandBuilders;
using Lookout.Options;
using Lookout.Sinks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OutlookServices;

/// <summary>
/// Provides the application entry point.
/// </summary>
internal class Program
{
    /// <summary>
    /// Runs the application.
    /// </summary>
    /// <param name="args">The command-line arguments.</param>
    /// <returns>The application exit code.</returns>
    public static async Task<int> Main(string[] args)
    {
        Environment.ExitCode = 0;
        Console.OutputEncoding = Encoding.UTF8;

        HostApplicationBuilder builder = Host.CreateApplicationBuilder(args);

        builder.Environment.ContentRootPath = AppContext.BaseDirectory;
        builder.Services
            .AddOptions<ExportAttachmentsOptions>()
            .BindConfiguration("ExportAttachments");

        builder.Services.AddSingleton<OutlookService>();
        builder.Services.AddSingleton<SinkFactory>();

        RootCommand rootCommand = RootCommandBuilder.Build(builder);

        return await rootCommand.Parse(args).InvokeAsync();
    }
}
