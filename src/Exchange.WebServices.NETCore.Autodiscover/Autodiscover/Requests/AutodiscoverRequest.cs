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

using System.IO.Compression;
using System.Net;
using System.Net.Http.Headers;
using System.Xml;

using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents the base class for all requested made to the Autodiscover service.
/// </summary>
internal abstract class AutodiscoverRequest
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverRequest" /> class.
    /// </summary>
    /// <param name="service">Autodiscover service associated with this request.</param>
    /// <param name="url">URL of Autodiscover service.</param>
    internal AutodiscoverRequest(AutodiscoverService service, Uri url)
    {
        Service = service;
        Url = url;
    }

    /// <summary>
    ///     Determines whether response is a redirection.
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    /// <returns>True if redirection response.</returns>
    internal static bool IsRedirectionResponse(IEwsHttpWebResponse httpWebResponse)
    {
        return httpWebResponse.StatusCode == HttpStatusCode.Redirect ||
               httpWebResponse.StatusCode == HttpStatusCode.Moved ||
               httpWebResponse.StatusCode == HttpStatusCode.RedirectKeepVerb ||
               httpWebResponse.StatusCode == HttpStatusCode.RedirectMethod;
    }

    /// <summary>
    ///     Validates the request.
    /// </summary>
    internal virtual void Validate()
    {
        Service.Validate();
    }

    /// <summary>
    ///     Executes this instance.
    /// </summary>
    /// <returns></returns>
    internal async Task<AutodiscoverResponse> InternalExecute()
    {
        Validate();

        try
        {
            var request = Service.PrepareHttpRequestMessageForUrl(Url);

            var needSignature = Service.Credentials != null && Service.Credentials.NeedSignature;
            var needTrace = Service.IsTraceEnabledFor(TraceFlags.AutodiscoverRequest);

            using (var memoryStream = new MemoryStream())
            {
                using (var writer = new EwsServiceXmlWriter(Service, memoryStream))
                {
                    writer.RequireWsSecurityUtilityNamespace = needSignature;
                    WriteSoapRequest(Url, writer);
                }

                if (needSignature)
                {
                    Service.Credentials.Sign(memoryStream);
                }

                if (needTrace)
                {
                    memoryStream.Position = 0;
                    Service.TraceXml(TraceFlags.AutodiscoverRequest, memoryStream);
                }

                request.Content = new ByteArrayContent(memoryStream.ToArray());
                request.Content.Headers.ContentType = new MediaTypeHeaderValue("text/xml")
                {
                    CharSet = "utf-8",
                };
            }

            using var client = Service.PrepareHttpClient();
            using var webResponse = new EwsHttpWebResponse(await client.SendAsync(request));

            if (IsRedirectionResponse(webResponse))
            {
                var response = CreateRedirectionResponse(webResponse);
                if (response != null)
                {
                    return response;
                }

                throw new ServiceRemoteException(Strings.InvalidRedirectionResponseReturned);
            }

            await using var responseStream = await GetResponseStream(webResponse);
            using (var memoryStream = new MemoryStream())
            {
                // Copy response stream to in-memory stream and reset to start
                await responseStream.CopyToAsync(memoryStream);

                memoryStream.Position = 0;

                Service.TraceResponse(webResponse, memoryStream);

                var ewsXmlReader = new EwsXmlReader(memoryStream);

                // WCF may not generate an XML declaration.
                await ewsXmlReader.ReadAsync().ConfigureAwait(false);

                if (ewsXmlReader.NodeType == XmlNodeType.XmlDeclaration)
                {
                    await ewsXmlReader.ReadStartElementAsync(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName)
                        .ConfigureAwait(false);
                }
                else if (ewsXmlReader.NodeType != XmlNodeType.Element ||
                         ewsXmlReader.LocalName != XmlElementNames.SOAPEnvelopeElementName ||
                         ewsXmlReader.NamespaceUri != EwsUtilities.GetNamespaceUri(XmlNamespace.Soap))
                {
                    throw new ServiceXmlDeserializationException(Strings.InvalidAutodiscoverServiceResponse);
                }

                ReadSoapHeaders(ewsXmlReader);

                var response = ReadSoapBody(ewsXmlReader);

                await ewsXmlReader.ReadEndElementAsync(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName)
                    .ConfigureAwait(false);

                if (response.ErrorCode == AutodiscoverErrorCode.NoError)
                {
                    return response;
                }

                throw new AutodiscoverResponseException(response.ErrorCode, response.ErrorMessage);
            }
        }
        catch (EwsHttpClientException ex)
        {
            if (ex.IsProtocolError && ex.Response != null)
            {
                var httpWebResponse = EwsHttpWebRequestFactory.CreateExceptionResponse(ex);

                if (IsRedirectionResponse(httpWebResponse))
                {
                    Service.ProcessHttpResponseHeaders(TraceFlags.AutodiscoverResponseHttpHeaders, httpWebResponse);

                    var response = CreateRedirectionResponse(httpWebResponse);
                    if (response != null)
                    {
                        return response;
                    }
                }
                else
                {
                    await ProcessEwsHttpClientException(ex);
                }
            }

            // Wrap exception if the above code block didn't throw
            throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
        }
        catch (XmlException ex)
        {
            Service.TraceMessage(TraceFlags.AutodiscoverConfiguration, $"XML parsing error: {ex.Message}");

            // Wrap exception
            throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
        }
        catch (IOException ex)
        {
            Service.TraceMessage(TraceFlags.AutodiscoverConfiguration, $"I/O error: {ex.Message}");

            // Wrap exception
            throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
        }
    }

    /// <summary>
    ///     Processes the web exception.
    /// </summary>
    /// <param name="webException">The web exception.</param>
    private async Task ProcessEwsHttpClientException(EwsHttpClientException webException)
    {
        if (webException.Response == null)
        {
            return;
        }

        var httpWebResponse = EwsHttpWebRequestFactory.CreateExceptionResponse(webException);

        if (httpWebResponse.StatusCode == HttpStatusCode.InternalServerError)
        {
            // If tracing is enabled, we read the entire response into a MemoryStream so that we
            // can pass it along to the ITraceListener. Then we parse the response from the 
            // MemoryStream.
            SoapFaultDetails soapFaultDetails;
            if (Service.IsTraceEnabledFor(TraceFlags.AutodiscoverRequest))
            {
                using var memoryStream = new MemoryStream();
                await using (var serviceResponseStream = await GetResponseStream(httpWebResponse))
                {
                    // Copy response to in-memory stream and reset position to start.
                    await serviceResponseStream.CopyToAsync(memoryStream);
                    memoryStream.Position = 0;
                }

                Service.TraceResponse(httpWebResponse, memoryStream);

                var reader = new EwsXmlReader(memoryStream);
                soapFaultDetails = ReadSoapFault(reader);
            }
            else
            {
                await using var stream = await GetResponseStream(httpWebResponse);
                var reader = new EwsXmlReader(stream);
                soapFaultDetails = ReadSoapFault(reader);
            }

            if (soapFaultDetails != null)
            {
                throw new ServiceResponseException(new ServiceResponse(soapFaultDetails));
            }
        }
        else
        {
            Service.ProcessHttpErrorResponse(httpWebResponse, webException);
        }
    }

    /// <summary>
    ///     Create a redirection response.
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    private AutodiscoverResponse? CreateRedirectionResponse(IEwsHttpWebResponse httpWebResponse)
    {
        var location = httpWebResponse.Headers.Location;
        if (location != null)
        {
            try
            {
                var redirectionUri = new Uri(Url, location);
                if (redirectionUri.Scheme == "http" || redirectionUri.Scheme == "https")
                {
                    var response = CreateServiceResponse();
                    response.ErrorCode = AutodiscoverErrorCode.RedirectUrl;
                    response.RedirectionUrl = redirectionUri;
                    return response;
                }

                Service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    $"Invalid redirection URL '{redirectionUri}' returned by Autodiscover service."
                );
            }
            catch (UriFormatException)
            {
                Service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    $"Invalid redirection location '{location}' returned by Autodiscover service."
                );
            }
        }
        else
        {
            Service.TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                "Redirection response returned by Autodiscover service without redirection location."
            );
        }

        return null;
    }

    /// <summary>
    ///     Reads the SOAP fault.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>SOAP fault details.</returns>
    private static SoapFaultDetails? ReadSoapFault(EwsXmlReader reader)
    {
        SoapFaultDetails? soapFaultDetails = null;

        try
        {
            // WCF may not generate an XML declaration.
            reader.Read();
            if (reader.NodeType == XmlNodeType.XmlDeclaration)
            {
                reader.Read();
            }

            if (!reader.IsStartElement() || reader.LocalName != XmlElementNames.SOAPEnvelopeElementName)
            {
                return soapFaultDetails;
            }

            // Get the namespace URI from the envelope element and use it for the rest of the parsing.
            // If it's not 1.1 or 1.2, we can't continue.
            var soapNamespace = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);
            if (soapNamespace == XmlNamespace.NotSpecified)
            {
                return soapFaultDetails;
            }

            reader.Read();

            // Skip SOAP header.
            if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPHeaderElementName))
            {
                do
                {
                    reader.Read();
                } while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPHeaderElementName));

                // Queue up the next read
                reader.Read();
            }

            // Parse the fault element contained within the SOAP body.
            if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPBodyElementName))
            {
                do
                {
                    reader.Read();

                    // Parse Fault element
                    if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPFaultElementName))
                    {
                        soapFaultDetails = SoapFaultDetails.Parse(reader, soapNamespace);
                    }
                } while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPBodyElementName));
            }

            reader.ReadEndElement(soapNamespace, XmlElementNames.SOAPEnvelopeElementName);
        }
        catch (XmlException)
        {
            // If response doesn't contain a valid SOAP fault, just ignore exception and
            // return null for SOAP fault details.
        }

        return soapFaultDetails;
    }

    /// <summary>
    ///     Writes the autodiscover SOAP request.
    /// </summary>
    /// <param name="requestUrl">Request URL.</param>
    /// <param name="writer">The writer.</param>
    internal void WriteSoapRequest(Uri requestUrl, EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
        writer.WriteAttributeValue(
            "xmlns",
            EwsUtilities.AutodiscoverSoapNamespacePrefix,
            EwsUtilities.AutodiscoverSoapNamespace
        );
        writer.WriteAttributeValue(
            "xmlns",
            EwsUtilities.WsAddressingNamespacePrefix,
            EwsUtilities.WsAddressingNamespace
        );
        writer.WriteAttributeValue(
            "xmlns",
            EwsUtilities.EwsXmlSchemaInstanceNamespacePrefix,
            EwsUtilities.EwsXmlSchemaInstanceNamespace
        );
        if (writer.RequireWsSecurityUtilityNamespace)
        {
            writer.WriteAttributeValue(
                "xmlns",
                EwsUtilities.WsSecurityUtilityNamespacePrefix,
                EwsUtilities.WsSecurityUtilityNamespace
            );
        }

        writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);

        Service.Credentials?.EmitExtraSoapHeaderNamespaceAliases(writer.InternalWriter);

        writer.WriteElementValue(
            XmlNamespace.Autodiscover,
            XmlElementNames.RequestedServerVersion,
            Service.RequestedServerVersion.ToString()
        );

        writer.WriteElementValue(XmlNamespace.WSAddressing, XmlElementNames.Action, GetWsAddressingActionName());

        writer.WriteElementValue(XmlNamespace.WSAddressing, XmlElementNames.To, requestUrl.AbsoluteUri);

        WriteExtraCustomSoapHeadersToXml(writer);

        Service.Credentials?.SerializeWSSecurityHeaders(writer.InternalWriter);

        Service.DoOnSerializeCustomSoapHeaders(writer.InternalWriter);

        writer.WriteEndElement(); // soap:Header

        writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

        WriteBodyToXml(writer);

        writer.WriteEndElement(); // soap:Body
        writer.WriteEndElement(); // soap:Envelope
        writer.Flush();
    }

    /// <summary>
    ///     Write extra headers.
    /// </summary>
    /// <param name="writer">The writer</param>
    internal virtual void WriteExtraCustomSoapHeadersToXml(EwsServiceXmlWriter writer)
    {
        // do nothing here. 
        // currently used only by GetUserSettingRequest to emit the BinarySecret header.
    }

    /// <summary>
    ///     Writes XML body.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteBodyToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Autodiscover, GetRequestXmlElementName());
        WriteAttributesToXml(writer);
        WriteElementsToXml(writer);

        writer.WriteEndElement(); // m:this.GetXmlElementName()
    }

    /// <summary>
    ///     Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
    /// </summary>
    /// <param name="response">HttpWebResponse.</param>
    /// <returns>ResponseStream</returns>
    protected static async Task<Stream> GetResponseStream(IEwsHttpWebResponse response)
    {
        var contentEncoding = response.ContentEncoding;
        var responseStream = await response.GetResponseStream();

        if (contentEncoding.ToLowerInvariant().Contains("gzip"))
        {
            return new GZipStream(responseStream, CompressionMode.Decompress);
        }

        if (contentEncoding.ToLowerInvariant().Contains("deflate"))
        {
            return new DeflateStream(responseStream, CompressionMode.Decompress);
        }

        return responseStream;
    }

    /// <summary>
    ///     Read SOAP headers.
    /// </summary>
    /// <param name="reader">EwsXmlReader</param>
    internal void ReadSoapHeaders(EwsXmlReader reader)
    {
        reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);
        do
        {
            reader.Read();

            ReadSoapHeader(reader);
        } while (!reader.IsEndElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName));
    }

    /// <summary>
    ///     Reads a single SOAP header.
    /// </summary>
    /// <param name="reader">EwsXmlReader</param>
    internal virtual void ReadSoapHeader(EwsXmlReader reader)
    {
        // Is this the ServerVersionInfo?
        if (reader.IsStartElement(XmlNamespace.Autodiscover, XmlElementNames.ServerVersionInfo))
        {
            Service.ServerInfo = ReadServerVersionInfo(reader);
        }
    }

    /// <summary>
    ///     Read ServerVersionInfo SOAP header.
    /// </summary>
    /// <param name="reader">EwsXmlReader</param>
    private static ExchangeServerInfo ReadServerVersionInfo(EwsXmlReader reader)
    {
        var serverInfo = new ExchangeServerInfo();
        do
        {
            reader.Read();

            if (reader.IsStartElement())
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.MajorVersion:
                    {
                        serverInfo.MajorVersion = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.MinorVersion:
                    {
                        serverInfo.MinorVersion = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.MajorBuildNumber:
                    {
                        serverInfo.MajorBuildNumber = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.MinorBuildNumber:
                    {
                        serverInfo.MinorBuildNumber = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.Version:
                    {
                        serverInfo.VersionString = reader.ReadElementValue();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.ServerVersionInfo));

        return serverInfo;
    }

    /// <summary>
    ///     Read SOAP body.
    /// </summary>
    /// <param name="reader">EwsXmlReader</param>
    internal AutodiscoverResponse ReadSoapBody(EwsXmlReader reader)
    {
        reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
        var responses = LoadFromXml(reader);
        reader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
        return responses;
    }

    /// <summary>
    ///     Loads responses from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns></returns>
    internal AutodiscoverResponse LoadFromXml(EwsXmlReader reader)
    {
        var elementName = GetResponseXmlElementName();
        reader.ReadStartElement(XmlNamespace.Autodiscover, elementName);
        var response = CreateServiceResponse();
        response.LoadFromXml(reader, elementName);
        return response;
    }

    /// <summary>
    ///     Gets the name of the request XML element.
    /// </summary>
    /// <returns></returns>
    internal abstract string GetRequestXmlElementName();

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns></returns>
    internal abstract string GetResponseXmlElementName();

    /// <summary>
    ///     Gets the WS-Addressing action name.
    /// </summary>
    /// <returns></returns>
    internal abstract string GetWsAddressingActionName();

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <returns>AutodiscoverResponse</returns>
    internal abstract AutodiscoverResponse CreateServiceResponse();

    /// <summary>
    ///     Writes attributes to request XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal abstract void WriteAttributesToXml(EwsServiceXmlWriter writer);

    /// <summary>
    ///     Writes elements to request XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal abstract void WriteElementsToXml(EwsServiceXmlWriter writer);

    /// <summary>
    ///     Gets the service.
    /// </summary>
    internal AutodiscoverService Service { get; }

    /// <summary>
    ///     Gets the URL.
    /// </summary>
    internal Uri Url { get; }
}
