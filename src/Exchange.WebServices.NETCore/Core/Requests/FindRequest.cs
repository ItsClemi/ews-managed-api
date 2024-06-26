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
///     Represents an abstract Find request.
/// </summary>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class FindRequest<TResponse> : MultiResponseServiceRequest<TResponse>
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the parent folder ids.
    /// </summary>
    public FolderIdWrapperList ParentFolderIds { get; } = new();

    /// <summary>
    ///     Gets or sets the search filter. Available search filter classes include SearchFilter.IsEqualTo,
    ///     SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection. If SearchFilter
    ///     is null, no search filters are applied.
    /// </summary>
    public SearchFilter? SearchFilter { get; set; }

    /// <summary>
    ///     Gets or sets the query string for indexed search.
    /// </summary>
    public string QueryString { get; set; }

    /// <summary>
    ///     Gets or sets the query string highlight terms.
    /// </summary>
    internal bool ReturnHighlightTerms { get; set; }

    /// <summary>
    ///     Gets or sets the view controlling the number of items or folders returned.
    /// </summary>
    public ViewBase View { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FindRequest&lt;TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal FindRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        View.InternalValidate(this);

        // query string parameter is only valid for Exchange2010 or higher
        //
        if (!string.IsNullOrEmpty(QueryString) && Service.RequestedServerVersion < ExchangeVersion.Exchange2010)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.ParameterIncompatibleWithRequestVersion,
                    "queryString",
                    ExchangeVersion.Exchange2010
                )
            );
        }

        // ReturnHighlightTerms parameter is only valid for Exchange2013 or higher
        //
        if (ReturnHighlightTerms && Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.ParameterIncompatibleWithRequestVersion,
                    "returnHighlightTerms",
                    ExchangeVersion.Exchange2013
                )
            );
        }

        // SeekToConditionItemView is only valid for Exchange2013 or higher
        //
        if (View is SeekToConditionItemView && Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.ParameterIncompatibleWithRequestVersion,
                    "SeekToConditionItemView",
                    ExchangeVersion.Exchange2013
                )
            );
        }

        if (!string.IsNullOrEmpty(QueryString) && SearchFilter != null)
        {
            throw new ServiceLocalException(Strings.BothSearchFilterAndQueryStringCannotBeSpecified);
        }
    }

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>XML element name.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return ParentFolderIds.Count;
    }

    /// <summary>
    ///     Gets the group by clause.
    /// </summary>
    /// <returns>The group by clause, null if the request does not have or support grouping.</returns>
    internal virtual Grouping GetGroupBy()
    {
        return null;
    }

    /// <summary>
    ///     Writes XML attributes.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        View.WriteAttributesToXml(writer);
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        View.WriteToXml(writer, GetGroupBy());

        if (SearchFilter != null)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Restriction);
            SearchFilter.WriteToXml(writer);
            writer.WriteEndElement(); // Restriction
        }

        View.WriteOrderByToXml(writer);

        ParentFolderIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.ParentFolderIds);

        if (!string.IsNullOrEmpty(QueryString))
        {
            // Emit the QueryString
            //
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.QueryString);

            if (ReturnHighlightTerms)
            {
                writer.WriteAttributeString(
                    XmlAttributeNames.ReturnHighlightTerms,
                    ReturnHighlightTerms.ToString().ToLowerInvariant()
                );
            }

            writer.WriteValue(QueryString, XmlElementNames.QueryString);
            writer.WriteEndElement();
        }
    }
}
