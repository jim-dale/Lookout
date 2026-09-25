namespace Lookout.Commands;

using System;
using System.Collections.Generic;
using Lookout.Options;
using Lookout.Sinks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OutlookServices;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Lists contacts from the Outlook contact folders.
/// </summary>
internal class ListContactsCommand : BackgroundService
{
    private readonly ListContactsOptions options;
    private readonly IOptions<ListContactsOptions> options1;
    private readonly OutlookService outlookService;
    private readonly SinkFactory formatterFactory;
    private readonly IHostApplicationLifetime hostApplicationLifetime;
    private readonly ILogger<ListContactsCommand> logger;
    private ContactSink? sink;

    /// <summary>
    /// Initializes a new instance of the <see cref="ListContactsCommand"/> class.
    /// </summary>
    /// <param name="options">The command options.</param>
    /// <param name="outlookService">The service used to access Outlook.</param>
    /// <param name="formatterFactory">The sink to use to output the contacts.</param>
    /// <param name="hostApplicationLifetime">The application lifetime used to stop the host.</param>
    /// <param name="logger">The logger used to record command activity.</param>
    public ListContactsCommand(IOptions<ListContactsOptions> options, OutlookService outlookService, SinkFactory formatterFactory, IHostApplicationLifetime hostApplicationLifetime, ILogger<ListContactsCommand> logger)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(options.Value);
        ArgumentNullException.ThrowIfNull(outlookService);
        ArgumentNullException.ThrowIfNull(hostApplicationLifetime);
        ArgumentNullException.ThrowIfNull(logger);

        this.options = options.Value;
        this.options1 = options;
        this.outlookService = outlookService;
        this.formatterFactory = formatterFactory;
        this.hostApplicationLifetime = hostApplicationLifetime;
        this.logger = logger;
    }

    /// <summary>
    /// Processes the contact folders in an Outlook store.
    /// </summary>
    /// <param name="store">The Outlook store to process.</param>
    internal void ProcessStore(Outlook.Store store)
    {
        ArgumentNullException.ThrowIfNull(store);

        this.logger.ProcessingStore(store.DisplayName, store.FilePath, store.StoreID, store.ExchangeStoreType);

        if (this.outlookService.TryGetRootFolder(store, out Outlook.MAPIFolder rootFolder))
        {
            this.ProcessFolder(rootFolder, this.options.MaxRecurseLevel, 0);
        }
    }

    /// <summary>
    /// Processes a contact folder and its eligible subfolders.
    /// </summary>
    /// <param name="folder">The folder to process.</param>
    /// <param name="maxLevel">The maximum folder recursion level.</param>
    /// <param name="level">The current folder recursion level.</param>
    internal void ProcessFolder(Outlook.MAPIFolder folder, int maxLevel, int level)
    {
        ArgumentNullException.ThrowIfNull(folder);

        this.logger.ProcessingFolder(folder.Name, folder.FullFolderPath, level);

        this.ProcessFolderItems(folder.Items);

        if (level < maxLevel && folder.Folders is not null)
        {
            foreach (Outlook.MAPIFolder subFolder in folder.Folders)
            {
                if (subFolder.DefaultItemType == Outlook.OlItemType.olContactItem)
                {
                    this.ProcessFolder(subFolder, maxLevel, level + 1);
                }
            }
        }
        else
        {
            this.logger.MaxRecusionLevelReached();
        }
    }

    /// <summary>
    /// Processes the contact items in a folder.
    /// </summary>
    /// <param name="items">The items to process.</param>
    internal void ProcessFolderItems(Outlook.Items items)
    {
        ArgumentNullException.ThrowIfNull(items);

        IEnumerable<Outlook.ContactItem> contacts = items.OfType<Outlook.ContactItem>()
            .OrderBy(c => c.FileAs);

        foreach (Outlook.ContactItem item in contacts)
        {
            this.ProcessContactItem(item);
        }
    }

    /// <summary>
    /// Adds a contact's configured properties to the output table.
    /// </summary>
    /// <param name="item">The contact item to process.</param>
    internal void ProcessContactItem(Outlook.ContactItem item)
    {
        this.sink?.Emit(item);
    }

    /// <inheritdoc />
    protected override Task ExecuteAsync(CancellationToken stoppingToken)
    {
        this.logger.ListContacts(this.options.MaxRecurseLevel);

        try
        {
            IEnumerable<Outlook.Store> item = this.outlookService.GetStores();

            this.sink = this.formatterFactory.CreateContactSink(this.options.PropertyNames);
            this.sink.Start();

            foreach (Outlook.Store store in item)
            {
                this.ProcessStore(store);
            }

            this.sink.Complete();
        }
        catch (Exception ex)
        {
            this.logger.ListContactsFailed(ex);
        }

        this.hostApplicationLifetime.StopApplication();
        return Task.CompletedTask;
    }
}
