namespace OutlookServices;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

#pragma warning disable SA1101

/// <summary>
/// Provides extension methods for Outlook interop objects.
/// </summary>
public static class OutlookExtensions
{
    extension(Store store)
    {
        /// <summary>
        /// Determines whether the store matches the specified descriptor.
        /// </summary>
        /// <param name="storeDescriptor">The display name, store ID, or file path fragment to find.</param>
        /// <returns><see langword="true"/> if the store matches the descriptor; otherwise, <see langword="false"/>.</returns>
        public bool IsDescriptorMatch(string storeDescriptor)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(storeDescriptor);

            bool result = false;

            if (!string.IsNullOrWhiteSpace(store.DisplayName) && store.DisplayName.Contains(storeDescriptor, StringComparison.CurrentCultureIgnoreCase))
            {
                result = true;
            }
            else if (!string.IsNullOrWhiteSpace(store.StoreID) && store.StoreID.Contains(storeDescriptor, StringComparison.CurrentCultureIgnoreCase))
            {
                result = true;
            }
            else if (!string.IsNullOrWhiteSpace(store.FilePath) && store.FilePath.Contains(storeDescriptor, StringComparison.CurrentCultureIgnoreCase))
            {
                result = true;
            }

            return result;
        }
    }

    extension(MailItem mailItem)
    {
        /// <summary>
        /// Gets the transport message headers from the mail item.
        /// </summary>
        /// <returns>The transport message headers, or <see langword="null"/> if no headers are available.</returns>
        public string? GetHeaders()
        {
            string? result = (string?)mailItem.PropertyAccessor.GetProperty(Constants.PR_TRANSPORT_MESSAGE_HEADERS);

            return result;
        }
    }

    extension(Attachment attachment)
    {
        /// <summary>
        /// Attempts to get the file name of the attachment.
        /// </summary>
        /// <param name="result">When this method returns, contains the attachment file name if one is available; otherwise, <see langword="null"/>.</param>
        /// <returns><see langword="true"/> if a non-empty file name was retrieved; otherwise, <see langword="false"/>.</returns>
        public bool TryGetFileName([NotNullWhen(true)] out string? result)
        {
            try
            {
                result = attachment.FileName;
            }
            catch (COMException)
            {
                result = default;
            }

            return !string.IsNullOrWhiteSpace(result);
        }
    }
}
