namespace Lookout.Sinks;

using System;
using System.Collections.Generic;
using OutlookServices;
using Spectre.Console;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Collects Outlook.ContactItem property values into an in-memory Table, adding columns for encountered property names
/// and appending a row for each emitted contact.
/// </summary>
/// <remarks>Columns are created on the first emission; subsequent emissions append rows. Property values are
/// converted to their string representation when added to the table. The sink is initialized with a set of property
/// names that guide which properties are extracted. Thread-safety is not guaranteed.</remarks>
internal class TableContactSink : ContactSink
{
    private readonly Table table = new();
    private readonly string[] propertyNames;
    private bool addColumns = true;

    /// <summary>
    /// Initializes a new instance of the <see cref="TableContactSink"/> class.
    /// </summary>
    /// <param name="propertyNames">The names of the contact properties to display.</param>
    internal TableContactSink(string[] propertyNames)
    {
        ArgumentNullException.ThrowIfNull(propertyNames);

        this.propertyNames = propertyNames;
    }

    /// <summary>
    /// Adds the contact's selected properties to the internal table as a new row, adding the columns before the first data row is added.
    /// </summary>
    /// <param name="item">The Outlook.ContactItem whose selected properties are written to the table.</param>
    internal override void Emit(Outlook.ContactItem item)
    {
        IReadOnlyDictionary<string, object?> properties = item.GetPropertyValues(this.propertyNames);

        if (this.addColumns)
        {
            foreach (var property in properties)
            {
                this.table.AddColumn(property.Key);
            }

            this.addColumns = false;
        }

        string[] values = properties.Select(p => $"{p.Value}").ToArray();
        this.table.AddRow(values);
    }

    /// <summary>
    /// Writes the formatted table to the console when the table contains one or more rows.
    /// </summary>
    /// <remarks>If the table is empty, no output is produced. Output is written using AnsiConsole.</remarks>
    internal override void Complete()
    {
        if (this.table.Rows.Count > 0)
        {
            AnsiConsole.Write(this.table);
        }
    }
}
