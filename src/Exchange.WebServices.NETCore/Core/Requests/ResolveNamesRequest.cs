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
///     Represents a ResolveNames request.
/// </summary>
internal sealed class ResolveNamesRequest : MultiResponseServiceRequest<ResolveNamesResponse>
{
    private static readonly IReadOnlyDictionary<ResolveNameSearchLocation, string> SearchScopeMap =
        new Dictionary<ResolveNameSearchLocation, string>
        {
            // @formatter:off
            { ResolveNameSearchLocation.DirectoryOnly, "ActiveDirectory" },
            { ResolveNameSearchLocation.DirectoryThenContacts, "ActiveDirectoryContacts" },
            { ResolveNameSearchLocation.ContactsOnly, "Contacts" },
            { ResolveNameSearchLocation.ContactsThenDirectory, "ContactsActiveDirectory" },
            // @formatter:on
        };

    /// <summary>
    ///     Gets or sets the name to resolve.
    /// </summary>
    /// <value>The name to resolve.</value>
    public string NameToResolve { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether to return full contact data or not.
    /// </summary>
    /// <value>
    ///     <c>true</c> if should return full contact data; otherwise, <c>false</c>.
    /// </value>
    public bool ReturnFullContactData { get; set; }

    /// <summary>
    ///     Gets or sets the search location.
    /// </summary>
    /// <value>The search scope.</value>
    public ResolveNameSearchLocation SearchLocation { get; set; }

    /// <summary>
    ///     Gets or sets the PropertySet for Contact Data
    /// </summary>
    /// <value>The PropertySet</value>
    public PropertySet? ContactDataPropertySet { get; set; }

    /// <summary>
    ///     Gets the parent folder ids.
    /// </summary>
    /// <value>The parent folder ids.</value>
    public FolderIdWrapperList ParentFolderIds { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="ResolveNamesRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal ResolveNamesRequest(ExchangeService service)
        : base(service, ServiceErrorHandling.ThrowOnError)
    {
    }

    /// <summary>
    ///     Asserts the valid.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateNonBlankStringParam(NameToResolve, "NameToResolve");
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Service response.</returns>
    protected override ResolveNamesResponse CreateServiceResponse(ExchangeService service, int responseIndex)
    {
        return new ResolveNamesResponse(service);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.ResolveNames;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.ResolveNamesResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.ResolveNamesResponseMessage;
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
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.ReturnFullContactData, ReturnFullContactData);

        SearchScopeMap.TryGetValue(SearchLocation, out var searchScope);

        EwsUtilities.Assert(
            !string.IsNullOrEmpty(searchScope),
            "ResolveNameRequest.WriteAttributesToXml",
            "The specified search location cannot be mapped to an EWS search scope."
        );

        string? propertySet = null;
        if (ContactDataPropertySet != null)
        {
            PropertySet.DefaultPropertySetMap.TryGetValue(ContactDataPropertySet.BasePropertySet, out propertySet);
        }

        if (!Service.Exchange2007CompatibilityMode)
        {
            writer.WriteAttributeValue(XmlAttributeNames.SearchScope, searchScope);
        }

        if (!string.IsNullOrEmpty(propertySet))
        {
            writer.WriteAttributeValue(XmlAttributeNames.ContactDataShape, propertySet);
        }
    }

    /// <summary>
    ///     Writes the elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        ParentFolderIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.ParentFolderIds);

        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.UnresolvedEntry, NameToResolve);
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
