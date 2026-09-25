namespace Lookout.Commands;

using System;
using System.Diagnostics.CodeAnalysis;
using Lookout.Options;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OutlookServices;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Exports attachments from the Outlook stores available to the current user.
/// </summary>
internal class ExportAttachmentsCommand : BackgroundService
{
    private readonly ExportAttachmentsOptions options;
    private readonly OutlookService outlookService;
    private readonly IHostApplicationLifetime hostApplicationLifetime;
    private readonly ILogger<ListStoresCommand> logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExportAttachmentsCommand"/> class.
    /// </summary>
    /// <param name="options">The attachment export options.</param>
    /// <param name="outlookService">The service used to access Outlook.</param>
    /// <param name="hostApplicationLifetime">The application lifetime used to stop the host.</param>
    /// <param name="logger">The logger used to record command activity.</param>
    public ExportAttachmentsCommand(IOptions<ExportAttachmentsOptions> options, OutlookService outlookService, IHostApplicationLifetime hostApplicationLifetime, ILogger<ListStoresCommand> logger)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(options.Value);
        ArgumentNullException.ThrowIfNull(outlookService);
        ArgumentNullException.ThrowIfNull(hostApplicationLifetime);
        ArgumentNullException.ThrowIfNull(logger);

        this.options = options.Value;
        this.outlookService = outlookService;
        this.hostApplicationLifetime = hostApplicationLifetime;
        this.logger = logger;
    }

    /// <summary>
    /// Processes the folders in an Outlook store.
    /// </summary>
    /// <param name="store">The Outlook store to process.</param>
    internal void ProcessStore(Outlook.Store store)
    {
        ArgumentNullException.ThrowIfNull(store);

        this.logger.ProcessingStore(store.DisplayName, store.FilePath, store.StoreID, store.ExchangeStoreType);

        if (this.outlookService.TryGetRootFolder(store, out Outlook.MAPIFolder rootFolder))
        {
            if (this.options.WhatIf == false)
            {
                this.options.TargetFolder.Create();
            }

            this.ProcessFolder(rootFolder, this.options.MaxRecurseLevel, 0);
        }
    }

    /// <summary>
    /// Processes a folder and its subfolders.
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
                this.ProcessFolder(subFolder, maxLevel, level + 1);
            }
        }
        else
        {
            this.logger.MaxRecusionLevelReached();
        }
    }

    /// <summary>
    /// Processes the mail and meeting items in a folder.
    /// </summary>
    /// <param name="items">The items to process.</param>
    internal void ProcessFolderItems(Outlook.Items items)
    {
        ArgumentNullException.ThrowIfNull(items);

        foreach (Outlook.MailItem item in items.OfType<Outlook.MailItem>())
        {
            this.ProcessMailItem(item);
        }

        foreach (Outlook.MeetingItem item in items.OfType<Outlook.MeetingItem>())
        {
            this.ProcessMeetingItem(item);
        }
    }

    /// <summary>
    /// Processes the attachments in a mail item.
    /// </summary>
    /// <param name="item">The mail item to process.</param>
    internal void ProcessMailItem(Outlook.MailItem item)
    {
        ArgumentNullException.ThrowIfNull(item);

        if (item.Attachments is not null && item.Attachments.Count > 0)
        {
            this.logger.MailAttachmentSummary(item.Subject, item.Attachments.Count);
            this.ProcessAttachments(item.Attachments);
        }
    }

    /// <summary>
    /// Processes the attachments in a meeting item.
    /// </summary>
    /// <param name="item">The meeting item to process.</param>
    internal void ProcessMeetingItem(Outlook.MeetingItem item)
    {
        ArgumentNullException.ThrowIfNull(item);

        if (item.Attachments is not null && item.Attachments.Count > 0)
        {
            this.logger.MeetingAttachmentSummary(item.Subject, item.Attachments.Count);
            this.ProcessAttachments(item.Attachments);
        }
    }

    /// <summary>
    /// Processes a collection of attachments.
    /// </summary>
    /// <param name="items">The attachments to process.</param>
    internal void ProcessAttachments(Outlook.Attachments? items)
    {
        if (items is not null)
        {
            foreach (Outlook.Attachment item in items)
            {
                this.ProcessAttachment(item);
            }
        }
    }

    /// <summary>
    /// Processes an attachment according to its type.
    /// </summary>
    /// <param name="item">The attachment to process.</param>
    internal void ProcessAttachment(Outlook.Attachment item)
    {
        Log.ProcessingAttachment(this.logger, item.DisplayName, item.Type, item.Size);

        switch (item.Type)
        {
            case Outlook.OlAttachmentType.olByValue:
                this.TrySaveAttachment(item, out _);
                break;
            case Outlook.OlAttachmentType.olEmbeddeditem:
                this.ProcessEmbeddedItemAttachment(item);
                break;
            case Outlook.OlAttachmentType.olOLE:
                this.TrySaveAttachment(item, out _);
                break;
            case Outlook.OlAttachmentType.olByReference:
            default:
                this.logger.UnsupportedAttachmentType(item.Type);
                break;
        }
    }

    /// <summary>
    /// Processes an embedded Outlook item attachment.
    /// </summary>
    /// <param name="item">The embedded item attachment to process.</param>
    internal void ProcessEmbeddedItemAttachment(Outlook.Attachment item)
    {
        try
        {
            // Try to save the embedded message
            if (this.TrySaveAttachment(item, out string? path))
            {
                if (this.options.WhatIf == false)
                {
                    using SharedItem sharedItem = this.outlookService.CreateSharedItemFrom(path);

                    if (sharedItem.TryCastAs(out Outlook.MailItem? mailItem))
                    {
                        this.ProcessMailItem(mailItem);
                    }
                    else if (sharedItem.TryCastAs(out Outlook.MeetingItem? meetingItem))
                    {
                        this.ProcessMeetingItem(meetingItem);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log exceptions and ignore because this is not a critical operation, i.e. it is best effort.
            this.logger.ProcessingSharedItemFailed(ex, item.DisplayName);
        }
    }

    /// <summary>
    /// Saves an attachment to the configured target folder.
    /// </summary>
    /// <param name="item">The attachment to save.</param>
    /// <param name="result">The path of the saved attachment, if successful.</param>
    /// <returns><see langword="true"/> if the attachment was saved or would be saved; otherwise, <see langword="false"/>.</returns>
    internal bool TrySaveAttachment(Outlook.Attachment item, [NotNullWhen(true)] out string? result)
    {
        result = default;

        if (item.Size > 0)
        {
            item.TryGetFileName(out string? destinationFileName);

            if (string.IsNullOrWhiteSpace(destinationFileName))
            {
                destinationFileName = $"{Guid.CreateVersion7():n}";
            }

            destinationFileName = FileHelpers.CleanFileName(destinationFileName);

            string destinationFilePath = FileHelpers.GetUniqueFilePath(this.options.TargetFolder, destinationFileName);

            destinationFilePath = Path.GetFullPath(destinationFilePath);

            Console.WriteLine($"Saving: \"{item.FileName}\" as \"{destinationFilePath}\"");
            this.logger.SavingAttachment(item.FileName, destinationFilePath);

            if (this.options.WhatIf == false)
            {
                item.SaveAsFile(destinationFilePath);
            }

            result = destinationFilePath;
        }

        return !string.IsNullOrWhiteSpace(result);
    }

    /// <inheritdoc />
    protected override Task ExecuteAsync(CancellationToken stoppingToken)
    {
        this.logger.ExportAttachments(this.options.TargetFolder, this.options.MaxRecurseLevel, this.options.WhatIf);

        try
        {
            IEnumerable<Outlook.Store> item = this.outlookService.GetStores();

            foreach (Outlook.Store store in item)
            {
                this.ProcessStore(store);
            }
        }
        catch (Exception ex)
        {
            this.logger.ExportAttachmentsFailed(ex);
        }

        this.hostApplicationLifetime.StopApplication();
        return Task.CompletedTask;
    }
}
