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
///     Represents the view settings in a folder search operation.
/// </summary>
[PublicAPI]
public sealed class PeopleIndexedItemView : PagedView
{
    /// <summary>
    ///     Gets the properties against which the returned items should be ordered.
    /// </summary>
    public OrderByCollection OrderBy { get; } = new();

    /// <summary>
    ///     Gets or sets the view filter.
    /// </summary>
    public ViewFilter? ViewFilter { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    /// <param name="offset">The offset of the view from the base point.</param>
    /// <param name="offsetBasePoint">The base point of the offset.</param>
    public PeopleIndexedItemView(int pageSize, int offset, OffsetBasePoint offsetBasePoint)
        : base(pageSize, offset, offsetBasePoint)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    public PeopleIndexedItemView(int pageSize)
        : base(pageSize)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemView" /> class.
    /// </summary>
    /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
    /// <param name="offset">The offset of the view from the base point.</param>
    public PeopleIndexedItemView(int pageSize, int offset)
        : base(pageSize, offset)
    {
        Offset = offset;
    }

    /// <summary>
    ///     Gets the type of service object this view applies to.
    /// </summary>
    /// <returns>A ServiceObjectType value.</returns>
    internal override ServiceObjectType GetServiceObjectType()
    {
        return ServiceObjectType.Persona;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        if (ViewFilter.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.ViewFilter, ViewFilter);
        }
    }

    /// <summary>
    ///     Gets the name of the view XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetViewXmlElementName()
    {
        return XmlElementNames.IndexedPageItemView;
    }

    /// <summary>
    ///     Validates this view.
    /// </summary>
    /// <param name="request">The request using this view.</param>
    internal override void InternalValidate(ServiceRequestBase request)
    {
        base.InternalValidate(request);

        if (ViewFilter.HasValue)
        {
            EwsUtilities.ValidateEnumVersionValue(ViewFilter, request.Service.RequestedServerVersion);
        }
    }

    /// <summary>
    ///     Writes OrderBy property to XML.
    /// </summary>
    /// <param name="writer">The writer</param>
    internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
    {
        OrderBy.WriteToXml(writer, XmlElementNames.SortOrder);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="groupBy">The group by clause.</param>
    internal override void WriteToXml(EwsServiceXmlWriter writer, Grouping groupBy)
    {
        writer.WriteStartElement(XmlNamespace.Messages, GetViewXmlElementName());

        InternalWriteViewToXml(writer);

        writer.WriteEndElement();
    }
}
