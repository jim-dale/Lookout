namespace Lookout.Sinks;

using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Provides a base sink for processing Outlook.ContactItem objects.
/// </summary>
/// <remarks>Override Start, Emit, and Complete to implement custom processing; base implementations are no-ops.</remarks>
internal class ContactSink
{
    /// <summary>
    /// Performs startup operations. The default implementation does nothing.
    /// </summary>
    internal virtual void Start()
    {
        // No-op
    }

    /// <summary>
    /// Emits a single ContactItem. The default implementation does nothing.
    /// </summary>
    /// <param name="item">The <see cref="Outlook.ContactItem"/> to emit.</param>
    internal virtual void Emit(Outlook.ContactItem item)
    {
        // No-op
    }

    /// <summary>
    /// Performs completion operations. The default implementation does nothing.
    /// </summary>
    internal virtual void Complete()
    {
        // No-op
    }
}
