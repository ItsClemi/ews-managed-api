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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a SyncFolderItems request.
/// </summary>
internal class SyncFolderItemsRequest : MultiResponseServiceRequest<SyncFolderItemsResponse>
{
    private int _maxChangesReturned = 100;
    private int _numberOfDays;

    /// <summary>
    ///     Gets or sets the property set.
    /// </summary>
    /// <value>The property set.</value>
    public PropertySet PropertySet { get; set; }

    /// <summary>
    ///     Gets or sets the sync folder id.
    /// </summary>
    /// <value>The sync folder id.</value>
    public FolderId SyncFolderId { get; set; }

    /// <summary>
    ///     Gets or sets the scope of the sync.
    /// </summary>
    /// <value>The scope of the sync.</value>
    public SyncFolderItemsScope SyncScope { get; set; }

    /// <summary>
    ///     Gets or sets the state of the sync.
    /// </summary>
    /// <value>The state of the sync.</value>
    public string SyncState { get; set; }

    /// <summary>
    ///     Gets the list of ignored item ids.
    /// </summary>
    /// <value>The ignored item ids.</value>
    public ItemIdWrapperList IgnoredItemIds { get; } = new();

    /// <summary>
    ///     Gets or sets the maximum number of changes returned by SyncFolderItems.
    ///     Values must be between 1 and 512.
    ///     Default is 100.
    /// </summary>
    public int MaxChangesReturned
    {
        get => _maxChangesReturned;

        set
        {
            if (value < 1 || value > 512)
            {
                throw new ArgumentException(Strings.MaxChangesMustBeBetween1And512);
            }

            _maxChangesReturned = value;
        }
    }

    /// <summary>
    ///     Gets or sets the number of days of content returned by SyncFolderItems.
    ///     Zero means return all content.
    ///     Default is zero.
    /// </summary>
    public int NumberOfDays
    {
        get => _numberOfDays;

        set
        {
            if (value < 0)
            {
                throw new ArgumentException(Strings.NumberOfDaysMustBePositive);
            }

            _numberOfDays = value;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SyncFolderItemsRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal SyncFolderItemsRequest(ExchangeService service)
        : base(service, ServiceErrorHandling.ThrowOnError)
    {
    }

    /// <summary>
    ///     Creates service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Service response.</returns>
    protected override SyncFolderItemsResponse CreateServiceResponse(ExchangeService service, int responseIndex)
    {
        return new SyncFolderItemsResponse(PropertySet);
    }

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of expected response messages.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return 1;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.SyncFolderItems;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.SyncFolderItemsResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.SyncFolderItemsResponseMessage;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateParam(PropertySet);
        EwsUtilities.ValidateParam(SyncFolderId);
        SyncFolderId.Validate(Service.RequestedServerVersion);

        // SyncFolderItemsScope enum was introduced with Exchange2010.  Only
        // value NormalItems is valid with previous server versions.
        if (Service.RequestedServerVersion < ExchangeVersion.Exchange2010 &&
            SyncScope != SyncFolderItemsScope.NormalItems)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.EnumValueIncompatibleWithRequestVersion,
                    SyncScope.ToString(),
                    SyncScope.GetType().Name,
                    ExchangeVersion.Exchange2010
                )
            );
        }

        // NumberOfDays was introduced with Exchange 2013.
        if (Service.RequestedServerVersion < ExchangeVersion.Exchange2013 && NumberOfDays != 0)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.ParameterIncompatibleWithRequestVersion,
                    "numberOfDays",
                    ExchangeVersion.Exchange2013
                )
            );
        }

        // SyncFolderItems can only handle summary properties
        PropertySet.ValidateForRequest(this, true /*summaryPropertiesOnly*/);
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        PropertySet.WriteToXml(writer, ServiceObjectType.Item);

        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SyncFolderId);
        SyncFolderId.WriteToXml(writer);
        writer.WriteEndElement();

        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SyncState, SyncState);

        IgnoredItemIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.Ignore);

        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MaxChangesReturned, MaxChangesReturned);

        if (Service.RequestedServerVersion >= ExchangeVersion.Exchange2010)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SyncScope, SyncScope);
        }

        if (NumberOfDays != 0)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.NumberOfDays, _numberOfDays);
        }
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }
}
