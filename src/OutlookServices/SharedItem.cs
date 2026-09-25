namespace OutlookServices;

using System.Diagnostics.CodeAnalysis;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Represents an Outlook item that has been opened using <c>Outlook.NameSpace.OpenSharedItem</c> and
/// provides a small, disposable wrapper around the raw COM object.
/// </summary>
/// <remarks>
/// This type stores the underlying COM object dynamically and exposes helper methods that perform
/// best-effort operations. Callers should dispose the instance when finished to ensure the underlying
/// Outlook COM object is closed. The class is not thread-safe and is intended for short-lived usage.
/// </remarks>
public sealed class SharedItem : IDisposable
{
    /// <summary>
    /// The underlying Outlook COM object returned from Outlook.NameSpace.OpenSharedItem(string).
    /// May be <c>null</c> when no item is opened or after <see cref="Dispose"/> has been called.
    /// </summary>
    private dynamic? obj;

    /// <summary>
    /// Attempts to open a shared Outlook item at the specified <paramref name="path"/> using the
    /// provided Outlook session.
    /// </summary>
    /// <param name="path">The filesystem or Outlook path of the shared item to open. Cannot be null or whitespace.</param>
    /// <param name="session">The Outlook session (<see cref="Outlook.NameSpace"/>) used to open the item. Cannot be <c>null</c>.</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is null or whitespace.</exception>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="session"/> is <c>null</c>.</exception>
    /// <exception cref="InvalidOperationException">Thrown when this instance already has an opened item and has not been disposed.</exception>
    public void OpenSharedItem(string path, Outlook.NameSpace session)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(session);

        // Throw an exception if this instance is being re-used without being disposed.
        if (this.obj is not null)
        {
            throw new InvalidOperationException();
        }

        this.obj = session.OpenSharedItem(path);
    }

    /// <summary>
    /// Attempts to cast the currently opened underlying item to the requested type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The reference type to attempt to cast the underlying item to (for example, an Outlook.MailItem).</typeparam>
    /// <param name="result">When this method returns, contains the cast result when the cast succeeded; otherwise <c>null</c>.</param>
    /// <returns><c>true</c> if the underlying item could be cast to <typeparamref name="T"/>; otherwise <c>false</c>.</returns>
    public bool TryCastAs<T>([NotNullWhen(true)] out T? result)
        where T : class
    {
        result = this.obj as T;
        return result is not null;
    }

    /// <summary>
    /// Closes and releases the underlying Outlook item if one is opened.
    /// </summary>
    /// <remarks>
    /// The underlying item's inspector (if any) is closed using <see cref="Outlook.OlInspectorClose.olDiscard"/>.
    /// After calling <see cref="Dispose"/>, this instance can be reused to open another item but that practice
    /// is discouraged; prefer creating a new <see cref="SharedItem"/> instance instead.
    /// </remarks>
    public void Dispose()
    {
        this.obj?.Close(Outlook.OlInspectorClose.olDiscard);
        this.obj = null;
    }
}
