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

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data.Credentials;


namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an abstract binding to an Exchange Service.
/// </summary>
[PublicAPI]
public abstract class ExchangeServiceBase : IDisposable
{
    /// <summary>
    ///     Default UserAgent
    /// </summary>
    private const string DefaultUserAgent = "ExchangeServicesClient/" + EwsUtilities.BuildVersion;

    /// <summary>
    ///     Special HTTP status code that indicates that the account is locked.
    /// </summary>
    internal const HttpStatusCode AccountIsLocked = (HttpStatusCode)456;

    private static readonly object LockObj = new();

    /// <summary>
    ///     The binary secret.
    /// </summary>
    private static byte[]? _binarySecret;

    /// <summary>
    ///     Underlying HttpWebRequest.
    /// </summary>
    private readonly HttpClient _httpClient;

    private readonly HttpClientHandler _httpClientHandler = new()
    {
        AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip,
    };


    private ExchangeCredentials? _credentials;
    private TimeZoneDefinition? _timeZoneDefinition;

    private bool _traceEnabled;
    private ITraceListener? _traceListener = new EwsTraceListener();
    private string _userAgent = DefaultUserAgent;


    /// <summary>
    ///     Gets or sets the cookie container.
    /// </summary>
    /// <value>The cookie container.</value>
    public CookieContainer CookieContainer
    {
        get => _httpClientHandler.CookieContainer;
        set => _httpClientHandler.CookieContainer = value;
    }

    /// <summary>
    ///     Gets the time zone this service is scoped to.
    /// </summary>
    internal TimeZoneInfo TimeZone { get; }

    /// <summary>
    ///     Gets a time zone definition generated from the time zone info to which this service is scoped.
    /// </summary>
    public TimeZoneDefinition TimeZoneDefinition => _timeZoneDefinition ??= new TimeZoneDefinition(TimeZone);

    /// <summary>
    ///     Gets or sets a value indicating whether client latency info is push to server.
    /// </summary>
    public bool SendClientLatencies { get; set; } = true;

    /// <summary>
    ///     Gets or sets a value indicating whether tracing is enabled.
    /// </summary>
    [MemberNotNullWhen(true, nameof(_traceListener))]
    public bool TraceEnabled
    {
        get => _traceEnabled;

        set
        {
            _traceEnabled = value;
            if (_traceEnabled && _traceListener == null)
            {
                _traceListener = new EwsTraceListener();
            }
        }
    }

    /// <summary>
    ///     Gets or sets the trace flags.
    /// </summary>
    /// <value>The trace flags.</value>
    public TraceFlags TraceFlags { get; set; } = TraceFlags.All;

    /// <summary>
    ///     Gets or sets the trace listener.
    /// </summary>
    /// <value>The trace listener.</value>
    public ITraceListener? TraceListener
    {
        get => _traceListener;

        set
        {
            _traceListener = value;
            _traceEnabled = value != null;
        }
    }

    /// <summary>
    ///     Gets or sets the credentials used to authenticate with the Exchange Web Services. Setting the Credentials property
    ///     automatically sets the UseDefaultCredentials to false.
    /// </summary>
    public ExchangeCredentials? Credentials
    {
        get => _credentials;

        set
        {
            _credentials = value;
            _httpClientHandler.UseDefaultCredentials = false;
            CookieContainer = new CookieContainer(); // Changing credentials resets the Cookie container

            // Set httpClient internal credentials
            ClientHandlerCredentials = value switch
            {
                DualAuthCredentials authCredentials => authCredentials.Credentials,
                OAuthCredentials authCredentials => authCredentials.Credentials,
                WebCredentials webCredentials => webCredentials.Credentials,
                _ => null,
            };
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
        get => _httpClientHandler.UseDefaultCredentials;

        set
        {
            _httpClientHandler.UseDefaultCredentials = value;

            if (value)
            {
                _credentials = null;
                CookieContainer = new CookieContainer(); // Changing credentials resets the Cookie container
            }
        }
    }

    /// <summary>
    ///     Gets or sets authentication information for the request.
    /// </summary>
    /// <returns>
    ///     An <see cref="T:System.Net.ICredentials" /> that contains the authentication credentials associated with the
    ///     request. The default is null.
    /// </returns>
    internal ICredentials? ClientHandlerCredentials
    {
        get => _httpClientHandler.Credentials;
        set => _httpClientHandler.Credentials = value;
    }


    /// <summary>
    ///     Gets or sets the timeout used when sending HTTP requests and when receiving HTTP responses, in milliseconds.
    ///     Defaults to 100000.
    /// </summary>
    public int Timeout
    {
        get => (int)_httpClient.Timeout.TotalMilliseconds;

        set
        {
            if (value < 1)
            {
                throw new ArgumentException(Strings.TimeoutMustBeGreaterThanZero);
            }

            _httpClient.Timeout = TimeSpan.FromMilliseconds(value);
        }
    }

    /// <summary>
    ///     Gets or sets a value that indicates whether HTTP pre-authentication should be performed.
    /// </summary>
    public bool PreAuthenticate
    {
        get => _httpClientHandler.PreAuthenticate;
        set => _httpClientHandler.PreAuthenticate = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether GZip compression encoding should be accepted.
    /// </summary>
    /// <remarks>
    ///     This value will tell the server that the client is able to handle GZip compression encoding. The server
    ///     will only send Gzip compressed content if it has been configured to do so.
    /// </remarks>
    public bool AcceptGzipEncoding { get; set; } = true;

    /// <summary>
    ///     Gets the requested server version.
    /// </summary>
    /// <value>The requested server version.</value>
    public ExchangeVersion RequestedServerVersion { get; } = ExchangeVersion.Exchange2013_SP1;

    /// <summary>
    ///     Gets or sets the user agent.
    /// </summary>
    /// <value>The user agent.</value>
    public string UserAgent
    {
        get => _userAgent;
        set => _userAgent = value + " (" + DefaultUserAgent + ")";
    }

    /// <summary>
    ///     Gets information associated with the server that processed the last request.
    ///     Will be null if no requests have been processed.
    /// </summary>
    public ExchangeServerInfo ServerInfo { get; internal set; }

    /// <summary>
    ///     Gets or sets the web proxy that should be used when sending requests to EWS.
    ///     Set this property to null to use the default web proxy.
    /// </summary>
    public IWebProxy? WebProxy
    {
        get => _httpClientHandler.Proxy;
        set => _httpClientHandler.Proxy = value;
    }

    /// <summary>
    ///     Gets or sets if the request to the internet resource should contain a Connection HTTP header with the value
    ///     Keep-alive
    /// </summary>
    public bool KeepAlive
    {
        get => !(_httpClient.DefaultRequestHeaders.ConnectionClose ?? false);
        set => _httpClient.DefaultRequestHeaders.ConnectionClose = !value;
    }

    /// <summary>
    ///     Gets or sets the name of the connection group for the request.
    /// </summary>
    [Obsolete("Has no effect")]
    public string ConnectionGroupName { get; set; }

    /// <summary>
    ///     Gets or sets the request id for the request.
    /// </summary>
    public string ClientRequestId { get; set; }

    /// <summary>
    ///     Gets or sets a flag to indicate whether the client requires the server side to return the  request id.
    /// </summary>
    public bool ReturnClientRequestId { get; set; }

    /// <summary>
    ///     Gets a collection of HTTP headers that will be sent with each request to EWS.
    /// </summary>
    public IDictionary<string, string> HttpHeaders { get; } = new Dictionary<string, string>();

    /// <summary>
    ///     Gets a collection of HTTP headers from the last response.
    /// </summary>
    public IDictionary<string, string> HttpResponseHeaders { get; } = new Dictionary<string, string>();

    /// <summary>
    ///     For testing: suppresses generation of the SOAP version header.
    /// </summary>
    internal bool SuppressXmlVersionHeader { get; set; }

    /// <summary>
    /// Optional client specified SSL certificate validation callback.
    /// </summary>
    public Func<HttpRequestMessage, X509Certificate2?, X509Chain?, SslPolicyErrors, bool>?
        ServerCertificateValidationCallback
    {
        get => _httpClientHandler.ServerCertificateCustomValidationCallback;
        set => _httpClientHandler.ServerCertificateCustomValidationCallback = value;
    }

    /// <summary>
    ///     Gets or sets a value that indicates whether the request should follow redirection responses.
    /// </summary>
    /// <returns>
    ///     True if the request should automatically follow redirection responses from the Internet resource; otherwise, false.
    ///     The default value is true.
    /// </returns>
    internal bool AllowAutoRedirect
    {
        get => _httpClientHandler.AllowAutoRedirect;
        set => _httpClientHandler.AllowAutoRedirect = value;
    }


    /// <summary>
    ///     Gets the session key.
    /// </summary>
    internal static byte[] SessionKey
    {
        get
        {
            // this has to be computed only once.
            lock (LockObj)
            {
                if (_binarySecret == null)
                {
                    var randomNumberGenerator = RandomNumberGenerator.Create();
                    _binarySecret = new byte[256 / 8];
                    randomNumberGenerator.GetBytes(_binarySecret);
                }

                return _binarySecret;
            }
        }
    }


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
        _httpClient = new HttpClient(_httpClientHandler);

        TimeZone = timeZone;
        UseDefaultCredentials = true;
        AllowAutoRedirect = true; // Always set to true in PrepareHttpRequestFromUrl
        CookieContainer = new CookieContainer();
        KeepAlive = true;
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
        RequestedServerVersion = requestedServerVersion;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class.
    /// </summary>
    /// <param name="service">The other service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    internal ExchangeServiceBase(ExchangeServiceBase service, ExchangeVersion requestedServerVersion)
        : this(requestedServerVersion)
    {
        UseDefaultCredentials = service.UseDefaultCredentials;
        _credentials = service.Credentials;
        _traceEnabled = service._traceEnabled;
        _traceListener = service._traceListener;
        TraceFlags = service.TraceFlags;
        Timeout = service.Timeout;
        PreAuthenticate = service.PreAuthenticate;
        _userAgent = service._userAgent;
        AcceptGzipEncoding = service.AcceptGzipEncoding;
        KeepAlive = service.KeepAlive;
        TimeZone = service.TimeZone;
        HttpHeaders = service.HttpHeaders;
        WebProxy = service.WebProxy;

        ConnectionGroupName = service.ConnectionGroupName;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeServiceBase" /> class from existing one.
    /// </summary>
    /// <param name="service">The other service.</param>
    internal ExchangeServiceBase(ExchangeServiceBase service)
        : this(service, service.RequestedServerVersion)
    {
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            _httpClient.Dispose();
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }


    /// <summary>
    ///     Occurs when the http response headers of a server call is captured.
    /// </summary>
    public event ResponseHeadersCapturedHandler? OnResponseHeadersCaptured;

    /// <summary>
    ///     Provides an event that applications can implement to emit custom SOAP headers in requests that are sent to
    ///     Exchange.
    /// </summary>
    public event CustomXmlSerializationDelegate? OnSerializeCustomSoapHeaders;


    #region Event handlers

    /// <summary>
    ///     Calls the custom SOAP header serialization event handlers, if defined.
    /// </summary>
    /// <param name="writer">The XmlWriter to which to write the custom SOAP headers.</param>
    internal void DoOnSerializeCustomSoapHeaders(XmlWriter writer)
    {
        EwsUtilities.Assert(writer != null, "ExchangeService.DoOnSerializeCustomSoapHeaders", "writer is null");

        OnSerializeCustomSoapHeaders?.Invoke(writer);
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


    #region Utilities

    /// <summary>
    ///     Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
    ///     based on the configuration of this service object.
    /// </summary>
    /// <param name="url">The URL that the HttpWebRequest should target.</param>
    /// <returns>A initialized instance of HttpWebRequest.</returns>
    internal async Task<EwsHttpWebRequest> PrepareHttpWebRequestForUrl(Uri url)
    {
        // Verify that the protocol is something that we can handle
        if (url.Scheme != "http" && url.Scheme != "https")
        {
            throw new ServiceLocalException(string.Format(Strings.UnsupportedWebProtocol, url.Scheme));
        }

        var request = new EwsHttpWebRequest(_httpClient, url)
        {
            ContentType = "text/xml; charset=utf-8",
            Accept = "text/xml",
            Method = "POST",
            UserAgent = UserAgent,
        };

        if (AcceptGzipEncoding)
        {
            request.Headers.AcceptEncoding.ParseAdd("gzip,deflate");
        }

        if (!string.IsNullOrEmpty(ClientRequestId))
        {
            request.Headers.TryAddWithoutValidation("client-request-id", ClientRequestId);

            if (ReturnClientRequestId)
            {
                request.Headers.TryAddWithoutValidation("return-client-request-id", "true");
            }
        }

        if (HttpHeaders.Count > 0)
        {
            foreach (var (key, value) in HttpHeaders)
            {
                request.Headers.TryAddWithoutValidation(key, value);
            }
        }

        if (!UseDefaultCredentials)
        {
            var serviceCredentials = Credentials;
            if (serviceCredentials == null)
            {
                throw new ServiceLocalException(Strings.CredentialsRequired);
            }

            // TODO: Temporary fix for authentication on Linux platform
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                serviceCredentials = AdjustLinuxAuthentication(url, serviceCredentials);
            }

            // Apply credentials to the request
            await serviceCredentials.PrepareWebRequest(request).ConfigureAwait(false);
        }

        lock (HttpResponseHeaders)
        {
            HttpResponseHeaders.Clear();
        }

        return request;
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

            Uri? accountUnlockUrl = null;
            if (Uri.IsWellFormedUriString(location, UriKind.Absolute))
            {
                accountUnlockUrl = new Uri(location);
            }

            TraceMessage(responseTraceFlag, $"Account is locked. Unlock URL is {accountUnlockUrl}");

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
    [MemberNotNullWhen(true, nameof(TraceListener))]
    internal bool IsTraceEnabledFor(TraceFlags traceFlags)
    {
        return TraceEnabled && (TraceFlags & traceFlags) != 0;
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
    internal void TraceHttpRequestHeaders(TraceFlags traceType, EwsHttpWebRequest request)
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
        lock (HttpResponseHeaders)
        {
            HttpResponseHeaders.Clear();

            foreach (var (key, value) in headers)
            {
                if (HttpResponseHeaders.TryGetValue(key, out var existingValue))
                {
                    HttpResponseHeaders[key] = existingValue + "," + string.Join(",", value);
                }
                else
                {
                    HttpResponseHeaders.Add(key, string.Join(",", value));
                }
            }
        }

        OnResponseHeadersCaptured?.Invoke(headers);
    }

    /// <summary>
    ///     Converts the universal date time string to local date time.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>DateTime</returns>
    internal DateTime? ConvertUniversalDateTimeStringToLocalDateTime(string? value)
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

        if (TimeZone.Equals(TimeZoneInfo.Utc))
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
    internal static DateTime? ConvertStartDateToUnspecifiedDateTime(string? value)
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
        var dateTime = value.Kind switch
        {
            DateTimeKind.Unspecified => EwsUtilities.ConvertTime(value, TimeZone, TimeZoneInfo.Utc),
            DateTimeKind.Local => EwsUtilities.ConvertTime(value, TimeZoneInfo.Local, TimeZoneInfo.Utc),
            // The date is already in UTC, no need to convert it.
            _ => value,
        };

        return dateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
    }


    /// <summary>
    ///     Sets the user agent to a custom value
    /// </summary>
    /// <param name="userAgent">User agent string to set on the service</param>
    internal void SetCustomUserAgent(string userAgent)
    {
        _userAgent = userAgent;
    }

    #endregion


    internal static ExchangeCredentials AdjustLinuxAuthentication(Uri url, ExchangeCredentials serviceCredentials)
    {
        if (serviceCredentials is not WebCredentials webCredentials)
        {
            // Nothing to adjust
            return serviceCredentials;
        }

        if (webCredentials.Credentials is NetworkCredential networkCredentials)
        {
            return new CredentialCache
            {
                // @formatter:off
                { url, "NTLM", networkCredentials },
                { url, "Digest", networkCredentials },
                { url, "Basic", networkCredentials },
                // @formatter:on
            };
        }

        return serviceCredentials;
    }
}
