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

using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Xml;
using System.Runtime.InteropServices;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an abstract binding to an Exchange Service.
/// </summary>
public abstract class ExchangeServiceBase
{
    #region Const members

    private static readonly object lockObj = new object();

    private readonly ExchangeVersion requestedServerVersion = ExchangeVersion.Exchange2013_SP1;

    /// <summary>
    ///     Special HTTP status code that indicates that the account is locked.
    /// </summary>
    internal const HttpStatusCode AccountIsLocked = (HttpStatusCode)456;

    /// <summary>
    ///     The binary secret.
    /// </summary>
    private static byte[] binarySecret;

    #endregion


    #region Static members

    /// <summary>
    ///     Default UserAgent
    /// </summary>
    private static readonly string defaultUserAgent = "ExchangeServicesClient/" + EwsUtilities.BuildVersion;

    #endregion


    #region Fields

    /// <summary>
    ///     Occurs when the http response headers of a server call is captured.
    /// </summary>
    public event ResponseHeadersCapturedHandler OnResponseHeadersCaptured;

    private ExchangeCredentials credentials;
    private bool useDefaultCredentials;
    private int timeout = 100000;
    private bool traceEnabled;
    private bool sendClientLatencies = true;
    private TraceFlags traceFlags = TraceFlags.All;
    private ITraceListener traceListener = new EwsTraceListener();
    private bool preAuthenticate;
    private string userAgent = defaultUserAgent;
    private bool acceptGzipEncoding = true;
    private bool keepAlive = true;
    private string connectionGroupName;
    private string clientRequestId;
    private bool returnClientRequestId;
    private CookieContainer cookieContainer = new CookieContainer();
    private readonly TimeZoneInfo timeZone;
    private TimeZoneDefinition timeZoneDefinition;
    private ExchangeServerInfo serverInfo;
    private IWebProxy webProxy;
    private readonly IDictionary<string, string> httpHeaders = new Dictionary<string, string>();
    private readonly IDictionary<string, string> httpResponseHeaders = new Dictionary<string, string>();
    private IEwsHttpWebRequestFactory ewsHttpWebRequestFactory = new EwsHttpWebRequestFactory();

    #endregion


    #region Event handlers

    /// <summary>
    ///     Calls the custom SOAP header serialization event handlers, if defined.
    /// </summary>
    /// <param name="writer">The XmlWriter to which to write the custom SOAP headers.</param>
    internal void DoOnSerializeCustomSoapHeaders(XmlWriter writer)
    {
        EwsUtilities.Assert(writer != null, "ExchangeService.DoOnSerializeCustomSoapHeaders", "writer is null");

        if (OnSerializeCustomSoapHeaders != null)
        {
            OnSerializeCustomSoapHeaders(writer);
        }
    }

    #endregion


    #region Utilities

    /// <summary>
    ///     Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
    ///     based on the configuration of this service object.
    /// </summary>
    /// <param name="url">The URL that the HttpWebRequest should target.</param>
    /// <param name="acceptGzipEncoding">If true, ask server for GZip compressed content.</param>
    /// <param name="allowAutoRedirect">If true, redirection responses will be automatically followed.</param>
    /// <returns>A initialized instance of HttpWebRequest.</returns>
    internal IEwsHttpWebRequest PrepareHttpWebRequestForUrl(Uri url, bool acceptGzipEncoding, bool allowAutoRedirect)
    {
        // Verify that the protocol is something that we can handle
        if ((url.Scheme != "http") && (url.Scheme != "https"))
        {
            throw new ServiceLocalException(string.Format(Strings.UnsupportedWebProtocol, url.Scheme));
        }

        var request = HttpWebRequestFactory.CreateRequest(url);
        try
        {
            request.PreAuthenticate = PreAuthenticate;
            request.Timeout = Timeout;
            SetContentType(request);
            request.Method = "POST";
            request.UserAgent = UserAgent;
            request.AllowAutoRedirect = allowAutoRedirect;
            request.CookieContainer = CookieContainer;
            request.KeepAlive = keepAlive;
            request.ConnectionGroupName = connectionGroupName;

            if (acceptGzipEncoding)
            {
                request.Headers.AcceptEncoding.ParseAdd("gzip,deflate");
            }

            if (!string.IsNullOrEmpty(clientRequestId))
            {
                request.Headers.TryAddWithoutValidation("client-request-id", clientRequestId);
                if (returnClientRequestId)
                {
                    request.Headers.TryAddWithoutValidation("return-client-request-id", "true");
                }
            }

            if (webProxy != null)
            {
                request.Proxy = webProxy;
            }

            if (HttpHeaders.Count > 0)
            {
                HttpHeaders.ForEach(kv => request.Headers.TryAddWithoutValidation(kv.Key, kv.Value));
            }

            request.UseDefaultCredentials = UseDefaultCredentials;
            if (!request.UseDefaultCredentials)
            {
                var serviceCredentials = Credentials;
                if (serviceCredentials == null)
                {
                    throw new ServiceLocalException(Strings.CredentialsRequired);
                }

                // Temporary fix for authentication on Linux platform
                if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    serviceCredentials = AdjustLinuxAuthentication(url, serviceCredentials);
                }

                // Make sure that credentials have been authenticated if required
                serviceCredentials.PreAuthenticate();

                // Apply credentials to the request
                serviceCredentials.PrepareWebRequest(request);
            }

            lock (httpResponseHeaders)
            {
                httpResponseHeaders.Clear();
            }

            return request;
        }
        catch (Exception)
        {
            request.Dispose();
            throw;
        }
    }

    internal ExchangeCredentials AdjustLinuxAuthentication(Uri url, ExchangeCredentials serviceCredentials)
    {
        if (!(serviceCredentials is WebCredentials))
            // Nothing to adjust
        {
            return serviceCredentials;
        }

        var networkCredentials = ((WebCredentials)serviceCredentials).Credentials as NetworkCredential;
        if (networkCredentials != null)
        {
            var credentialCache = new CredentialCache();
            credentialCache.Add(url, "NTLM", networkCredentials);
            credentialCache.Add(url, "Digest", networkCredentials);
            credentialCache.Add(url, "Basic", networkCredentials);

            serviceCredentials = credentialCache;
        }

        return serviceCredentials;
    }

    internal virtual void SetContentType(IEwsHttpWebRequest request)
    {
        request.ContentType = "text/xml; charset=utf-8";
        request.Accept = "text/xml";
    }

    /// <summary>
    ///     Processes an HTTP error response
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    /// <param name="webException">The web exception.</param>
    /// <param name="responseHeadersTraceFlag">The trace flag for response headers.</param>
    /// <param name="responseTraceFlag">The trace flag for responses.</param>
    /// <remarks>
    ///     This method doesn't handle 500 ISE errors. This is handled by the caller since
    ///     500 ISE typically indicates that a SOAP fault has occurred and the handling of
    ///     a SOAP fault is currently service specific.
    /// </remarks>
    internal void InternalProcessHttpErrorResponse(
        IEwsHttpWebResponse httpWebResponse,
        EwsHttpClientException webException,
        TraceFlags responseHeadersTraceFlag,
        TraceFlags responseTraceFlag
    )
    {
        EwsUtilities.Assert(
            httpWebResponse.StatusCode != HttpStatusCode.InternalServerError,
            "ExchangeServiceBase.InternalProcessHttpErrorResponse",
            "InternalProcessHttpErrorResponse does not handle 500 ISE errors, the caller is supposed to handle this."
        );

        ProcessHttpResponseHeaders(responseHeadersTraceFlag, httpWebResponse);

        // Deal with new HTTP error code indicating that account is locked.
        // The "unlock" URL is returned as the status description in the response.
        if (httpWebResponse.StatusCode == AccountIsLocked)
        {
            var location = httpWebResponse.StatusDescription;

            Uri accountUnlockUrl = null;
            if (Uri.IsWellFormedUriString(location, UriKind.Absolute))
            {
                accountUnlockUrl = new Uri(location);
            }

            TraceMessage(responseTraceFlag, string.Format("Account is locked. Unlock URL is {0}", accountUnlockUrl));

            throw new AccountIsLockedException(
                string.Format(Strings.AccountIsLocked, accountUnlockUrl),
                accountUnlockUrl,
                webException
            );
        }
    }

    /// <summary>
    ///     Processes an HTTP error response.
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    /// <param name="webException">The web exception.</param>
    internal abstract void ProcessHttpErrorResponse(
        IEwsHttpWebResponse httpWebResponse,
        EwsHttpClientException webException
    );

    /// <summary>
    ///     Determines whether tracing is enabled for specified trace flag(s).
    /// </summary>
    /// <param name="traceFlags">The trace flags.</param>
    /// <returns>
    ///     True if tracing is enabled for specified trace flag(s).
    /// </returns>
    internal bool IsTraceEnabledFor(TraceFlags traceFlags)
    {
        return TraceEnabled && ((TraceFlags & traceFlags) != 0);
    }

    /// <summary>
    ///     Logs the specified string to the TraceListener if tracing is enabled.
    /// </summary>
    /// <param name="traceType">Kind of trace entry.</param>
    /// <param name="logEntry">The entry to log.</param>
    internal void TraceMessage(TraceFlags traceType, string logEntry)
    {
        if (IsTraceEnabledFor(traceType))
        {
            var traceTypeStr = traceType.ToString();
            var logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, logEntry);
            TraceListener.Trace(traceTypeStr, logMessage);
        }
    }

    /// <summary>
    ///     Logs the specified XML to the TraceListener if tracing is enabled.
    /// </summary>
    /// <param name="traceType">Kind of trace entry.</param>
    /// <param name="stream">The stream containing XML.</param>
    internal void TraceXml(TraceFlags traceType, MemoryStream stream)
    {
        if (IsTraceEnabledFor(traceType))
        {
            var traceTypeStr = traceType.ToString();
            var logMessage = EwsUtilities.FormatLogMessageWithXmlContent(traceTypeStr, stream);
            TraceListener.Trace(traceTypeStr, logMessage);
        }
    }

    /// <summary>
    ///     Traces the HTTP request headers.
    /// </summary>
    /// <param name="traceType">Kind of trace entry.</param>
    /// <param name="request">The request.</param>
    internal void TraceHttpRequestHeaders(TraceFlags traceType, IEwsHttpWebRequest request)
    {
        if (IsTraceEnabledFor(traceType))
        {
            var traceTypeStr = traceType.ToString();
            var headersAsString = EwsUtilities.FormatHttpRequestHeaders(request);
            var logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, headersAsString);
            TraceListener.Trace(traceTypeStr, logMessage);
        }
    }

    /// <summary>
    ///     Traces the HTTP response headers.
    /// </summary>
    /// <param name="traceType">Kind of trace entry.</param>
    /// <param name="response">The response.</param>
    internal void ProcessHttpResponseHeaders(TraceFlags traceType, IEwsHttpWebResponse response)
    {
        TraceHttpResponseHeaders(traceType, response);

        SaveHttpResponseHeaders(response.Headers);
    }

    /// <summary>
    ///     Traces the HTTP response headers.
    /// </summary>
    /// <param name="traceType">Kind of trace entry.</param>
    /// <param name="response">The response.</param>
    private void TraceHttpResponseHeaders(TraceFlags traceType, IEwsHttpWebResponse response)
    {
        if (IsTraceEnabledFor(traceType))
        {
            var traceTypeStr = traceType.ToString();
            var headersAsString = EwsUtilities.FormatHttpResponseHeaders(response);
            var logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, headersAsString);
            TraceListener.Trace(traceTypeStr, logMessage);
        }
    }

    /// <summary>
    ///     Save the HTTP response headers.
    /// </summary>
    /// <param name="headers">The response headers</param>
    private void SaveHttpResponseHeaders(HttpResponseHeaders headers)
    {
        lock (httpResponseHeaders)
        {
            httpResponseHeaders.Clear();

            foreach (var item in headers)
            {
                var key = item.Key;
                string existingValue;

                if (httpResponseHeaders.TryGetValue(key, out existingValue))
                {
                    httpResponseHeaders[key] = existingValue + "," + string.Join(",", item.Value);
                }
                else
                {
                    httpResponseHeaders.Add(key, string.Join(",", item.Value));
                }
            }
        }

        if (OnResponseHeadersCaptured != null)
        {
            OnResponseHeadersCaptured(headers);
        }
    }

    /// <summary>
    ///     Converts the universal date time string to local date time.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>DateTime</returns>
    internal DateTime? ConvertUniversalDateTimeStringToLocalDateTime(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        // Assume an unbiased date/time is in UTC. Convert to UTC otherwise.
        var dateTime = DateTime.Parse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal
        );

        if (TimeZone == TimeZoneInfo.Utc)
        {
            // This returns a DateTime with Kind.Utc
            return dateTime;
        }

        var localTime = EwsUtilities.ConvertTime(dateTime, TimeZoneInfo.Utc, TimeZone);

        if (EwsUtilities.IsLocalTimeZone(TimeZone))
        {
            // This returns a DateTime with Kind.Local
            return new DateTime(localTime.Ticks, DateTimeKind.Local);
        }

        // This returns a DateTime with Kind.Unspecified
        return localTime;
    }

    /// <summary>
    ///     Converts xs:dateTime string with either "Z", "-00:00" bias, or "" suffixes to
    ///     unspecified StartDate value ignoring the suffix.
    /// </summary>
    /// <param name="value">The string value to parse.</param>
    /// <returns>The parsed DateTime value.</returns>
    internal DateTime? ConvertStartDateToUnspecifiedDateTime(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        var dateTimeOffset = DateTimeOffset.Parse(value, CultureInfo.InvariantCulture);

        // Return only the date part with the kind==Unspecified.
        return dateTimeOffset.Date;
    }

    /// <summary>
    ///     Converts the date time to universal date time string.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>String representation of DateTime.</returns>
    internal string ConvertDateTimeToUniversalDateTimeString(DateTime value)
    {
        DateTime dateTime;

        switch (value.Kind)
        {
            case DateTimeKind.Unspecified:
                dateTime = EwsUtilities.ConvertTime(value, TimeZone, TimeZoneInfo.Utc);

                break;
            case DateTimeKind.Local:
                dateTime = EwsUtilities.ConvertTime(value, TimeZoneInfo.Local, TimeZoneInfo.Utc);

                break;
            default:
                // The date is already in UTC, no need to convert it.
                dateTime = value;

                break;
        }

        return dateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Register the custom auth module to support non-ascii upn authentication if the server supports that
    /// </summary>
    internal void RegisterCustomBasicAuthModule()
    {
        if (RequestedServerVersion >= ExchangeVersion.Exchange2013_SP1)
        {
            //BasicAuthModuleForUTF8.InstantiateIfNeeded();
        }
    }

    /// <summary>
    ///     Sets the user agent to a custom value
    /// </summary>
    /// <param name="userAgent">User agent string to set on the service</param>
    internal void SetCustomUserAgent(string userAgent)
    {
        this.userAgent = userAgent;
    }

    #endregion


    #region Constructors

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    internal ExchangeServiceBase()
        : this(TimeZoneInfo.Local)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    /// <param name="timeZone">The time zone to which the service is scoped.</param>
    internal ExchangeServiceBase(TimeZoneInfo timeZone)
    {
        this.timeZone = timeZone;
        UseDefaultCredentials = true;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    /// <param name="requestedServerVersion">The requested server version.</param>
    internal ExchangeServiceBase(ExchangeVersion requestedServerVersion)
        : this(requestedServerVersion, TimeZoneInfo.Local)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    /// <param name="requestedServerVersion">The requested server version.</param>
    /// <param name="timeZone">The time zone to which the service is scoped.</param>
    internal ExchangeServiceBase(ExchangeVersion requestedServerVersion, TimeZoneInfo timeZone)
        : this(timeZone)
    {
        this.requestedServerVersion = requestedServerVersion;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    /// <param name="service">The other service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    internal ExchangeServiceBase(ExchangeServiceBase service, ExchangeVersion requestedServerVersion)
        : this(requestedServerVersion)
    {
        useDefaultCredentials = service.useDefaultCredentials;
        credentials = service.credentials;
        traceEnabled = service.traceEnabled;
        traceListener = service.traceListener;
        traceFlags = service.traceFlags;
        timeout = service.timeout;
        preAuthenticate = service.preAuthenticate;
        userAgent = service.userAgent;
        acceptGzipEncoding = service.acceptGzipEncoding;
        keepAlive = service.keepAlive;
        connectionGroupName = service.connectionGroupName;
        timeZone = service.timeZone;
        httpHeaders = service.httpHeaders;
        ewsHttpWebRequestFactory = service.ewsHttpWebRequestFactory;
        webProxy = service.webProxy;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class from existing one.
    /// </summary>
    /// <param name="service">The other service.</param>
    internal ExchangeServiceBase(ExchangeServiceBase service)
        : this(service, service.RequestedServerVersion)
    {
    }

    #endregion


    #region Validation

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal virtual void Validate()
    {
    }

    #endregion


    #region Properties

    /// <summary>
    ///     Gets or sets the cookie container.
    /// </summary>
    /// <value>The cookie container.</value>
    public CookieContainer CookieContainer
    {
        get => cookieContainer;
        set => cookieContainer = value;
    }

    /// <summary>
    ///     Gets the time zone this service is scoped to.
    /// </summary>
    internal TimeZoneInfo TimeZone => timeZone;

    /// <summary>
    ///     Gets a time zone definition generated from the time zone info to which this service is scoped.
    /// </summary>
    public TimeZoneDefinition TimeZoneDefinition
    {
        get
        {
            if (timeZoneDefinition == null)
            {
                timeZoneDefinition = new TimeZoneDefinition(TimeZone);
            }

            return timeZoneDefinition;
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether client latency info is push to server.
    /// </summary>
    public bool SendClientLatencies
    {
        get => sendClientLatencies;

        set => sendClientLatencies = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether tracing is enabled.
    /// </summary>
    public bool TraceEnabled
    {
        get => traceEnabled;

        set
        {
            traceEnabled = value;
            if (traceEnabled && (traceListener == null))
            {
                traceListener = new EwsTraceListener();
            }
        }
    }

    /// <summary>
    ///     Gets or sets the trace flags.
    /// </summary>
    /// <value>The trace flags.</value>
    public TraceFlags TraceFlags
    {
        get => traceFlags;

        set => traceFlags = value;
    }

    /// <summary>
    ///     Gets or sets the trace listener.
    /// </summary>
    /// <value>The trace listener.</value>
    public ITraceListener TraceListener
    {
        get => traceListener;

        set
        {
            traceListener = value;
            traceEnabled = value != null;
        }
    }

    /// <summary>
    ///     Gets or sets the credentials used to authenticate with the Exchange Web Services. Setting the Credentials property
    ///     automatically sets the UseDefaultCredentials to false.
    /// </summary>
    public ExchangeCredentials Credentials
    {
        get => credentials;

        set
        {
            credentials = value;
            useDefaultCredentials = false;
            cookieContainer = new CookieContainer(); // Changing credentials resets the Cookie container
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the credentials of the user currently logged into Windows should be used to
    ///     authenticate with the Exchange Web Services. Setting UseDefaultCredentials to true automatically sets the
    ///     Credentials
    ///     property to null.
    /// </summary>
    public bool UseDefaultCredentials
    {
        get => useDefaultCredentials;

        set
        {
            useDefaultCredentials = value;

            if (value)
            {
                credentials = null;
                cookieContainer = new CookieContainer(); // Changing credentials resets the Cookie container
            }
        }
    }

    /// <summary>
    ///     Gets or sets the timeout used when sending HTTP requests and when receiving HTTP responses, in milliseconds.
    ///     Defaults to 100000.
    /// </summary>
    public int Timeout
    {
        get => timeout;

        set
        {
            if (value < 1)
            {
                throw new ArgumentException(Strings.TimeoutMustBeGreaterThanZero);
            }

            timeout = value;
        }
    }

    /// <summary>
    ///     Gets or sets a value that indicates whether HTTP pre-authentication should be performed.
    /// </summary>
    public bool PreAuthenticate
    {
        get => preAuthenticate;
        set => preAuthenticate = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether GZip compression encoding should be accepted.
    /// </summary>
    /// <remarks>
    ///     This value will tell the server that the client is able to handle GZip compression encoding. The server
    ///     will only send Gzip compressed content if it has been configured to do so.
    /// </remarks>
    public bool AcceptGzipEncoding
    {
        get => acceptGzipEncoding;
        set => acceptGzipEncoding = value;
    }

    /// <summary>
    ///     Gets the requested server version.
    /// </summary>
    /// <value>The requested server version.</value>
    public ExchangeVersion RequestedServerVersion => requestedServerVersion;

    /// <summary>
    ///     Gets or sets the user agent.
    /// </summary>
    /// <value>The user agent.</value>
    public string UserAgent
    {
        get => userAgent;
        set => userAgent = value + " (" + defaultUserAgent + ")";
    }

    /// <summary>
    ///     Gets information associated with the server that processed the last request.
    ///     Will be null if no requests have been processed.
    /// </summary>
    public ExchangeServerInfo ServerInfo
    {
        get => serverInfo;
        internal set => serverInfo = value;
    }

    /// <summary>
    ///     Gets or sets the web proxy that should be used when sending requests to EWS.
    ///     Set this property to null to use the default web proxy.
    /// </summary>
    public IWebProxy WebProxy
    {
        get => webProxy;
        set => webProxy = value;
    }

    /// <summary>
    ///     Gets or sets if the request to the internet resource should contain a Connection HTTP header with the value
    ///     Keep-alive
    /// </summary>
    public bool KeepAlive
    {
        get => keepAlive;

        set => keepAlive = value;
    }

    /// <summary>
    ///     Gets or sets the name of the connection group for the request.
    /// </summary>
    public string ConnectionGroupName
    {
        get => connectionGroupName;

        set => connectionGroupName = value;
    }

    /// <summary>
    ///     Gets or sets the request id for the request.
    /// </summary>
    public string ClientRequestId
    {
        get => clientRequestId;
        set => clientRequestId = value;
    }

    /// <summary>
    ///     Gets or sets a flag to indicate whether the client requires the server side to return the  request id.
    /// </summary>
    public bool ReturnClientRequestId
    {
        get => returnClientRequestId;
        set => returnClientRequestId = value;
    }

    /// <summary>
    ///     Gets a collection of HTTP headers that will be sent with each request to EWS.
    /// </summary>
    public IDictionary<string, string> HttpHeaders => httpHeaders;

    /// <summary>
    ///     Gets a collection of HTTP headers from the last response.
    /// </summary>
    public IDictionary<string, string> HttpResponseHeaders => httpResponseHeaders;

    /// <summary>
    ///     Gets the session key.
    /// </summary>
    internal static byte[] SessionKey
    {
        get
        {
            // this has to be computed only once.
            lock (lockObj)
            {
                if (binarySecret == null)
                {
                    var randomNumberGenerator = RandomNumberGenerator.Create();
                    binarySecret = new byte[256 / 8];
                    randomNumberGenerator.GetBytes(binarySecret);
                }

                return binarySecret;
            }
        }
    }

    /// <summary>
    ///     Gets or sets the HTTP web request factory.
    /// </summary>
    internal IEwsHttpWebRequestFactory HttpWebRequestFactory
    {
        get => ewsHttpWebRequestFactory;

        set =>
            // If new value is null, reset to default factory.
            ewsHttpWebRequestFactory = (value == null) ? new EwsHttpWebRequestFactory() : value;
    }

    /// <summary>
    ///     For testing: suppresses generation of the SOAP version header.
    /// </summary>
    internal bool SuppressXmlVersionHeader { get; set; }

    #endregion


    #region Events

    /// <summary>
    ///     Provides an event that applications can implement to emit custom SOAP headers in requests that are sent to
    ///     Exchange.
    /// </summary>
    public event CustomXmlSerializationDelegate OnSerializeCustomSoapHeaders;

    #endregion
}
