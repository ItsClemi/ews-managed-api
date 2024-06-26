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
///     Represents a GetUserAvailability request.
/// </summary>
internal sealed class GetUserAvailabilityRequest : SimpleServiceRequestBase
{
    /// <summary>
    ///     Gets a value indicating whether the TimeZoneContext SOAP header should be emitted.
    /// </summary>
    /// <value><c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.</value>
    internal override bool EmitTimeZoneHeader => true;

    /// <summary>
    ///     Gets a value indicating whether free/busy data is requested.
    /// </summary>
    internal bool IsFreeBusyViewRequested =>
        RequestedData == AvailabilityData.FreeBusy || RequestedData == AvailabilityData.FreeBusyAndSuggestions;

    /// <summary>
    ///     Gets a value indicating whether suggestions are requested.
    /// </summary>
    internal bool IsSuggestionsViewRequested =>
        RequestedData == AvailabilityData.Suggestions || RequestedData == AvailabilityData.FreeBusyAndSuggestions;

    /// <summary>
    ///     Gets or sets the attendees.
    /// </summary>
    public IEnumerable<AttendeeInfo> Attendees { get; set; }

    /// <summary>
    ///     Gets or sets the time window in which to retrieve user availability information.
    /// </summary>
    public TimeWindow TimeWindow { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating what data is requested (free/busy and/or suggestions).
    /// </summary>
    public AvailabilityData RequestedData { get; set; } = AvailabilityData.FreeBusyAndSuggestions;

    /// <summary>
    ///     Gets an object that allows you to specify options controlling the information returned
    ///     by the GetUserAvailability request.
    /// </summary>
    public AvailabilityOptions Options { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetUserAvailabilityRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal GetUserAvailabilityRequest(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.GetUserAvailabilityRequest;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        Options.Validate(TimeWindow.Duration);
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        // Only serialize the TimeZone property against an Exchange 2007 SP1 server.
        // Against Exchange 2010, the time zone is emitted in the request's SOAP header.
        if (writer.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
        {
            var legacyTimeZone = new LegacyAvailabilityTimeZone(writer.Service.TimeZone);

            legacyTimeZone.WriteToXml(writer, XmlElementNames.TimeZone);
        }

        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.MailboxDataArray);

        foreach (var attendee in Attendees)
        {
            attendee.WriteToXml(writer);
        }

        writer.WriteEndElement(); // MailboxDataArray

        Options.WriteToXml(writer, this);
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.GetUserAvailabilityResponse;
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var serviceResponse = new GetUserAvailabilityResults();

        if (IsFreeBusyViewRequested)
        {
            serviceResponse.AttendeesAvailability = new ServiceResponseCollection<AttendeeAvailability>();

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.FreeBusyResponseArray);

            do
            {
                reader.Read();

                if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FreeBusyResponse))
                {
                    var freeBusyResponse = new AttendeeAvailability();

                    freeBusyResponse.LoadFromXml(reader, XmlElementNames.ResponseMessage);

                    if (freeBusyResponse.ErrorCode == ServiceError.NoError)
                    {
                        freeBusyResponse.LoadFreeBusyViewFromXml(reader, Options.RequestedFreeBusyView);
                    }

                    serviceResponse.AttendeesAvailability.Add(freeBusyResponse);
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.FreeBusyResponseArray));
        }

        if (IsSuggestionsViewRequested)
        {
            serviceResponse.SuggestionsResponse = new SuggestionsResponse();

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.SuggestionsResponse);

            serviceResponse.SuggestionsResponse.LoadFromXml(reader, XmlElementNames.ResponseMessage);

            if (serviceResponse.SuggestionsResponse.ErrorCode == ServiceError.NoError)
            {
                serviceResponse.SuggestionsResponse.LoadSuggestedDaysFromXml(reader);
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.SuggestionsResponse);
        }

        return serviceResponse;
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response.</returns>
    internal async Task<GetUserAvailabilityResults> Execute(CancellationToken token)
    {
        return await InternalExecuteAsync<GetUserAvailabilityResults>(token).ConfigureAwait(false);
    }
}
