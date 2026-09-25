namespace ListContactsSpike;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using OutlookServices;
using OutlookServices.Options;
using Outlook = Microsoft.Office.Interop.Outlook;

#pragma warning disable

public static partial class StringExtensions
{
    extension(string? str)
    {
        public string EscapeCsvField()
        {
            if (str is null)
            {
                return string.Empty;
            }

            bool mustQuote = str.Contains(",") || str.Contains("\"") || str.Contains("\n") || str.Contains("\r");
            if (mustQuote)
            {
                // Escape double quotes by doubling them (RFC 4180)
                str = str.Replace("\"", "\"\"");
                return $"\"{str}\"";
            }

            return str;
        }
    }
}

public static partial class FileExtensions
{
    extension(File)
    {
        public static void WriteAllAsCsv<T>(string path, IEnumerable<T> items, string[] propertyNames, Converter<T, string> converter)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            ArgumentNullException.ThrowIfNull(items);
            ArgumentNullException.ThrowIfNull(propertyNames);
            ArgumentNullException.ThrowIfNull(converter);

            using TextWriter writer = new StreamWriter(path);

            File.WriteAllAsCsv(writer, items, propertyNames, converter);
        }

        public static void WriteAllAsCsv<T>(TextWriter writer, IEnumerable<T> items, string[] propertyNames, Converter<T, string> converter)
        {
            ArgumentNullException.ThrowIfNull(writer);
            ArgumentNullException.ThrowIfNull(items);
            ArgumentNullException.ThrowIfNull(propertyNames);
            ArgumentNullException.ThrowIfNull(converter);

            bool writeHeader = true;

            foreach (T item in items)
            {
                if (writeHeader)
                {
                    string csvHeader = string.Join(',', propertyNames);
                    writer.WriteLine(csvHeader);
                    writeHeader = false;
                }

                string line = converter.Invoke(item);
                writer.WriteLine(line);
            }
        }
    }
}

public class OutlookContactToCsvConverter
{
    private string[] propertyNames;

    public OutlookContactToCsvConverter(string[] propertyNames)
    {
        this.propertyNames = propertyNames;

        this.Converter = new(PropertyValuesToCsvRow);
    }

    public Converter<Outlook.ContactItem, string> Converter { get; }

    internal string PropertyValuesToCsvRow(Outlook.ContactItem item)
    {
        IReadOnlyDictionary<string, object?> properties = item.GetPropertyValues(this.propertyNames);

        return string.Join(',', properties.Select(p => $"{p.Value}".EscapeCsvField()));
    }
}

internal static class IEnumerableExtensions
{
    extension(IEnumerable<Outlook.ContactItem> items)
    {
        public string ToJsonArray(string[] propertyNames, JsonWriterOptions options)
        {
            using MemoryStream ms = new();
            using Utf8JsonWriter writer = new(ms, options);

            items.ToJsonArray(propertyNames, writer);

            return MemoryStreamToString(ms);
        }

        public void ToJsonArray(string[] propertyNames, Utf8JsonWriter writer)
        {
            writer.WriteStartArray();

            foreach (Outlook.ContactItem item in items)
            {
                var properties = item.GetPropertyValues(propertyNames);

                JsonSerializer.Serialize(writer, properties);
            }

            writer.WriteEndArray();
            writer.Flush();
        }
    }

    /// <summary>
    /// Converts a MemoryStream to a string using the specified encoding.
    /// </summary>
    /// <param name="ms">The MemoryStream to read from.</param>
    /// <param name="encoding">The text encoding to use.</param>
    /// <returns>The string representation of the MemoryStream's contents.</returns>
    internal static string MemoryStreamToString(MemoryStream ms, Encoding? encoding = null)
    {
        ArgumentNullException.ThrowIfNull(ms);

        encoding ??= Encoding.UTF8;

        // Ensure we read from the beginning
        if (ms.CanSeek)
        {
            ms.Position = 0;
        }

        using StreamReader reader = new(ms, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        {
            return reader.ReadToEnd();
        }
    }
}

internal class Program
{
    internal static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;

        string[] propertyNames = "FullName,MobileTelephoneNumber,HomeAddress".Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        IEnumerable<Outlook.ContactItem> contacts = GetFromOutlook().OrderBy(c => c.FirstName);

        OutlookContactToCsvConverter converter = new(propertyNames);

        File.WriteAllAsCsv(Console.Out, contacts, propertyNames, converter.Converter);
    }

    internal static IEnumerable<Outlook.ContactItem> GetFromOutlook()
    {
        using OutlookService outlookService = new(Options.Create(new OutlookServiceOptions()));

        foreach (Outlook.Store store in outlookService.GetAllStores())
        {
            var root = store.GetRootFolder();
            ////Outlook.MAPIFolder contactsFolder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            foreach (Outlook.Folder subFolder in root.Folders.OfType<Outlook.Folder>())
            {
                if (subFolder.DefaultItemType == Outlook.OlItemType.olContactItem)
                {
                    foreach (Outlook.ContactItem result in subFolder.Items.OfType<Outlook.ContactItem>())
                    {
                        yield return result;
                    }
                }
            }
        }
    }


    internal static string[] GetTelephoneNumbers(Outlook.ContactItem item)
    {
        string[] result =
        [
            NormalizePhoneNumber(item.AssistantTelephoneNumber),
            NormalizePhoneNumber(item.Business2TelephoneNumber),
            NormalizePhoneNumber(item.BusinessFaxNumber),
            NormalizePhoneNumber(item.BusinessTelephoneNumber),
            NormalizePhoneNumber(item.CallbackTelephoneNumber),
            NormalizePhoneNumber(item.CarTelephoneNumber),
            NormalizePhoneNumber(item.CompanyMainTelephoneNumber),
            NormalizePhoneNumber(item.Home2TelephoneNumber),
            NormalizePhoneNumber(item.HomeFaxNumber),
            NormalizePhoneNumber(item.HomeTelephoneNumber),
            NormalizePhoneNumber(item.MobileTelephoneNumber),
            NormalizePhoneNumber(item.OtherFaxNumber),
            NormalizePhoneNumber(item.OtherTelephoneNumber),
            NormalizePhoneNumber(item.PrimaryTelephoneNumber),
            NormalizePhoneNumber(item.RadioTelephoneNumber),
        ];

        result = result.Distinct().ToArray();

        return result.Where(i => string.IsNullOrWhiteSpace(i) == false).OrderBy(i => i).ToArray();
    }

    internal static string NormalizePhoneNumber(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return string.Empty;
        }

        string result = s.Replace("+", string.Empty, StringComparison.Ordinal)
            .Replace("-", string.Empty, StringComparison.Ordinal)
            .Replace("x", string.Empty, StringComparison.Ordinal)
            .Replace(";", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("(", string.Empty, StringComparison.Ordinal)
            .Replace(")", string.Empty, StringComparison.Ordinal);

        if (result.StartsWith("44"))
        {
            result = result[2..];
        }
        if (result.StartsWith('0'))
        {
            result = result[1..];
        }

        return result;
    }

    /// <summary>
    /// Normalize the specified string by applying Unicode normalization (Form C), removing non-word characters and
    /// diacritical marks, and converting letters to upper-case invariant.
    /// </summary>
    /// <remarks>Uses string.Normalize() (Form C), Regex.Replace with pattern [\W\p{M}] to remove non-word
    /// characters and combining marks, and ToUpperInvariant for culture-independent casing.</remarks>
    /// <param name="s">The input string to normalize. Null or whitespace yields an empty string.</param>
    /// <returns>The normalized string with Unicode normalization applied, non-word characters and combining marks removed, and
    /// letters converted to upper-case invariant; or an empty string if the input is null or whitespace.</returns>
    internal static string NormalizeString(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return string.Empty;
        }

        string result = Regex.Replace(s.Normalize(), @"[\W\p{M}]", string.Empty)
            .ToUpperInvariant();

        return result;
    }
}
