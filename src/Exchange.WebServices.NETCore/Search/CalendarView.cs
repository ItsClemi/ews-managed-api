/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a date range view of appointments in calendar folder search operations.
/// </summary>
[PublicAPI]
public sealed class CalendarView : ViewBase
{
    private int? _maxItemsReturned;

    /// <summary>
    ///     Gets or sets the start date.
    /// </summary>
    public DateTime StartDate { get; set; }

    /// <summary>
    ///     Gets or sets the end date.
    /// </summary>
    public DateTime EndDate { get; set; }

    /// <summary>
    ///     The maximum number of items the search operation should return.
    /// </summary>
    public int? MaxItemsReturned
    {
        get => _maxItemsReturned;

        set
        {
            if (value <= 0)
            {
                throw new ArgumentException(Strings.ValueMustBeGreaterThanZero, nameof(value));
            }

            _maxItemsReturned = value;
        }
    }

    /// <summary>
    ///     Gets or sets the search traversal mode. Defaults to ItemTraversal.Shallow.
    /// </summary>
    public ItemTraversal Traversal { get; set; }

    /// <summary>
    ///     Initializes a new instance of CalendarView.
    /// </summary>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    public CalendarView(DateTime startDate, DateTime endDate)
    {
        StartDate = startDate;
        EndDate = endDate;
    }

    /// <summary>
    ///     Initializes a new instance of CalendarView.
    /// </summary>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <param name="maxItemsReturned">The maximum number of items the search operation should return.</param>
    public CalendarView(DateTime startDate, DateTime endDate, int maxItemsReturned)
        : this(startDate, endDate)
    {
        MaxItemsReturned = maxItemsReturned;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Traversal, Traversal);
    }

    /// <summary>
    ///     Writes the search settings to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="groupBy">The group by clause.</param>
    internal override void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy)
    {
        // No search settings for calendar views.
    }

    /// <summary>
    ///     Writes OrderBy property to XML.
    /// </summary>
    /// <param name="writer">The writer</param>
    internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
    {
        // No OrderBy for calendar views.
    }

    /// <summary>
    ///     Gets the type of service object this view applies to.
    /// </summary>
    /// <returns>A ServiceObjectType value.</returns>
    internal override ServiceObjectType GetServiceObjectType()
    {
        return ServiceObjectType.Item;
    }

    /// <summary>
    ///     Validate instance.
    /// </summary>
    /// <param name="request">The request using this view.</param>
    internal override void InternalValidate(ServiceRequestBase request)
    {
        base.InternalValidate(request);

        if (EndDate < StartDate)
        {
            throw new ServiceValidationException(Strings.EndDateMustBeGreaterThanStartDate);
        }
    }

    /// <summary>
    ///     Write to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void InternalWriteViewToXml(EwsServiceXmlWriter writer)
    {
        base.InternalWriteViewToXml(writer);

        writer.WriteAttributeValue(XmlAttributeNames.StartDate, StartDate);
        writer.WriteAttributeValue(XmlAttributeNames.EndDate, EndDate);
    }

    /// <summary>
    ///     Gets the name of the view XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetViewXmlElementName()
    {
        return XmlElementNames.CalendarView;
    }

    /// <summary>
    ///     Gets the maximum number of items or folders the search operation should return.
    /// </summary>
    /// <returns>
    ///     The maximum number of items the search operation should return.
    /// </returns>
    internal override int? GetMaxEntriesReturned()
    {
        return MaxItemsReturned;
    }
}
