namespace OutlookServices;

#pragma warning disable SA1310 // Field names should not contain underscore, matches Media SDK

/// <summary>
/// Provides constants used when interacting with Outlook MAPI properties and item types.
/// </summary>
public static class Constants
{
    /// <summary>
    /// Identifies the Outlook MAPI namespace.
    /// </summary>
    public const string MapiNamespace = "MAPI";

    /// <summary>
    /// Identifies a standard Outlook note item.
    /// </summary>
    public const string Note = "IPM.Note";

    /// <summary>
    /// Identifies a multipart signed Outlook note item.
    /// </summary>
    public const string MultipartSignedNote = "IPM.Note.SMIME.MultipartSigned";

    /// <summary>
    /// Specifies the prefix used to construct MAPI property tag URLs.
    /// </summary>
    public const string PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x";

    /// <summary>
    /// Identifies the MAPI Boolean property type.
    /// </summary>
    public const string PT_BOOLEAN = "000B";

    /// <summary>
    /// Identifies the MAPI 8-bit string property type.
    /// </summary>
    public const string PT_STRING8 = "001E";

    /// <summary>
    /// Identifies the MAPI Unicode string property type.
    /// </summary>
    public const string PT_UNICODE = "001F";

    /// <summary>
    /// Identifies the transport message headers property.
    /// </summary>
    public const string PR_TRANSPORT_MESSAGE_HEADERS = PROPTAG + "007D" + PT_STRING8;

    /// <summary>
    /// Identifies the attachment content ID property.
    /// </summary>
    public const string PR_ATTACH_CONTENT_ID = PROPTAG + "3712" + PT_STRING8;

    /// <summary>
    /// Identifies the attachment hidden property.
    /// </summary>
    public const string PR_ATTACH_HIDDEN = PROPTAG + "7FFE" + PT_BOOLEAN;

    /// <summary>
    /// Represents the empty MAPI date value.
    /// </summary>
    public static readonly DateTime EmptyDate = new(4501, 1, 1, 0, 0, 0);
}
