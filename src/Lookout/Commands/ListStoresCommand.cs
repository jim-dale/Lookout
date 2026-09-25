namespace Lookout.Commands;

using System;
using Microsoft.Extensions.Logging;
using OutlookServices;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Lists the Outlook stores available to the current user.
/// </summary>
internal class ListStoresCommand : BackgroundService
{
    private readonly OutlookService outlookService;
    private readonly IHostApplicationLifetime hostApplicationLifetime;
    private readonly ILogger<ListStoresCommand> logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="ListStoresCommand"/> class.
    /// </summary>
    /// <param name="outlookService">The service used to access Outlook.</param>
    /// <param name="hostApplicationLifetime">The application lifetime used to stop the host.</param>
    /// <param name="logger">The logger used to record command activity.</param>
    public ListStoresCommand(OutlookService outlookService, IHostApplicationLifetime hostApplicationLifetime, ILogger<ListStoresCommand> logger)
    {
        ArgumentNullException.ThrowIfNull(outlookService);
        ArgumentNullException.ThrowIfNull(hostApplicationLifetime);
        ArgumentNullException.ThrowIfNull(logger);

        this.outlookService = outlookService;
        this.hostApplicationLifetime = hostApplicationLifetime;
        this.logger = logger;
    }

    /// <inheritdoc />
    protected override Task ExecuteAsync(CancellationToken stoppingToken)
    {
        this.logger.ListStores();

        try
        {
            foreach (Outlook.Store store in this.outlookService.GetAllStores())
            {
                Console.WriteLine($"Name=\"{store.DisplayName}\", Path=\"{store.FilePath}\", Id={store.StoreID}, Type={store.ExchangeStoreType}");
            }
        }
        catch (Exception ex)
        {
            this.logger.ListStoresFailed(ex);
        }

        this.hostApplicationLifetime.StopApplication();
        return Task.CompletedTask;
    }
}
