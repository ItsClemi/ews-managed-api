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

using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Enumeration of reasons that a hanging request may disconnect.
/// </summary>
internal enum HangingRequestDisconnectReason
{
    /// <summary>The server cleanly closed the connection.</summary>
    Clean,

    /// <summary>The client closed the connection.</summary>
    UserInitiated,

    /// <summary>The connection timed out do to a lack of a heartbeat received.</summary>
    Timeout,

    /// <summary>An exception occurred on the connection</summary>
    Exception,
}

/// <summary>
///     Represents a collection of arguments for the HangingServiceRequestBase.HangingRequestDisconnectHandler
///     delegate method.
/// </summary>
internal class HangingRequestDisconnectEventArgs : EventArgs
{
    /// <summary>
    ///     Gets the reason that the user was disconnected.
    /// </summary>
    public HangingRequestDisconnectReason Reason { get; internal set; }

    /// <summary>
    ///     Gets the exception that caused the disconnection. Can be null.
    /// </summary>
    public Exception? Exception { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="HangingRequestDisconnectEventArgs" /> class.
    /// </summary>
    /// <param name="reason">The reason.</param>
    /// <param name="exception">The exception.</param>
    internal HangingRequestDisconnectEventArgs(HangingRequestDisconnectReason reason, Exception? exception)
    {
        Reason = reason;
        Exception = exception;
    }
}

/// <summary>
///     Represents an abstract, hanging service request.
/// </summary>
internal abstract class HangingServiceRequestBase : ServiceRequestBase
{
    /// <summary>
    ///     Test switch to log all bytes that come across the wire.
    ///     Helpful when parsing fails before certain bytes hit the trace logs.
    /// </summary>
    internal const bool LogAllWireBytes = false;

    /// <summary>
    ///     Expected minimum frequency in responses, in milliseconds.
    /// </summary>
    protected readonly int _heartbeatFrequencyMilliseconds;

    /// <summary>
    ///     lock object
    /// </summary>
    private readonly object _lockObject = new();

    /// <summary>
    ///     Callback delegate to handle response objects
    /// </summary>
    private readonly HandleResponseObject _responseHandler;

    private readonly CancellationTokenSource _tokenSource = new();

    private System.Threading.Tasks.Task _readTask;

    /// <summary>
    ///     Request to the server.
    /// </summary>
    private EwsHttpWebRequest _request;

    /// <summary>
    ///     Response from the server.
    /// </summary>
    private IEwsHttpWebResponse _response;

    /// <summary>
    ///     Gets a value indicating whether this instance is connected.
    /// </summary>
    /// <value><c>true</c> if this instance is connected; otherwise, <c>false</c>.</value>
    internal bool IsConnected { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="HangingServiceRequestBase" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="handler">Callback delegate to handle response objects</param>
    /// <param name="heartbeatFrequency">Frequency at which we expect heartbeats, in milliseconds.</param>
    internal HangingServiceRequestBase(ExchangeService service, HandleResponseObject handler, int heartbeatFrequency)
        : base(service)
    {
        _responseHandler = handler;
        _heartbeatFrequencyMilliseconds = heartbeatFrequency;
    }

    /// <summary>
    ///     Occurs when the hanging request is disconnected.
    /// </summary>
    internal event HangingRequestDisconnectHandler OnDisconnect;

    /// <summary>
    ///     Executes the request.
    /// </summary>
    internal void InternalExecute()
    {
        lock (_lockObject)
        {
            var (request, response) =
                ValidateAndEmitRequest(headersOnly: true, _tokenSource.Token).GetAwaiter().GetResult();

            _request = request;
            _response = response;

            if (!IsConnected)
            {
                IsConnected = true;

                // Trace Http headers
                Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, _response);

                // Run parser on separate task-thread
                _readTask = System.Threading.Tasks.Task.Factory.StartNew(
                        async () => await ParseResponses(_tokenSource.Token).ConfigureAwait(false),
                        _tokenSource.Token,
                        TaskCreationOptions.LongRunning,
                        TaskScheduler.Default
                    )
                    .GetAwaiter()
                    .GetResult();
            }
        }
    }

    /// <summary>
    ///     Parses the responses.
    /// </summary>
    private async System.Threading.Tasks.Task ParseResponses(CancellationToken cancellationToken)
    {
        try
        {
            MemoryStream? responseCopy = null;

            try
            {
                var traceEwsResponse = Service.IsTraceEnabledFor(TraceFlags.EwsResponse);

                var responseStream = await _response.GetResponseStream(cancellationToken).ConfigureAwait(false);
                await using (responseStream.ConfigureAwait(false))
                {
                    var tracingStream = new HangingTraceStream(responseStream, Service)
                    {
                        ReadTimeout = 2 * _heartbeatFrequencyMilliseconds,
                    };

                    // EwsServiceMultiResponseXmlReader.Create causes a read.
                    if (traceEwsResponse)
                    {
                        responseCopy = new MemoryStream();
                        tracingStream.SetResponseCopy(responseCopy);
                    }

                    var ewsXmlReader = EwsServiceMultiResponseXmlReader.Create(tracingStream, Service);

                    while (IsConnected)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        object? responseObject;
                        if (traceEwsResponse)
                        {
                            try
                            {
                                responseObject = await ReadResponseAsync(ewsXmlReader, _response.Headers)
                                    .ConfigureAwait(false);
                            }
                            finally
                            {
                                Service.TraceXml(TraceFlags.EwsResponse, responseCopy);
                            }

                            // reset the stream collector.
                            await responseCopy.DisposeAsync().ConfigureAwait(false);
                            responseCopy = new MemoryStream();
                            tracingStream.SetResponseCopy(responseCopy);
                        }
                        else
                        {
                            responseObject = await ReadResponseAsync(ewsXmlReader, _response.Headers)
                                .ConfigureAwait(false);
                        }

                        _responseHandler(responseObject);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Operation was cancelled, do nothing
            }
            catch (TimeoutException ex)
            {
                // The connection timed out.
                Disconnect(HangingRequestDisconnectReason.Timeout, ex);
            }
            catch (IOException ex)
            {
                // Stream is closed, so disconnect.
                Disconnect(HangingRequestDisconnectReason.Exception, ex);
            }
            catch (EwsHttpClientException ex)
            {
                // Stream is closed, so disconnect.
                Disconnect(HangingRequestDisconnectReason.Exception, ex);
            }
            catch (ObjectDisposedException ex)
            {
                // Stream is closed, so disconnect.
                Disconnect(HangingRequestDisconnectReason.Exception, ex);
            }
            catch (NotSupportedException)
            {
                // This is thrown if we close the stream during a read operation due to a user method call.
                // Trying to delay closing until the read finishes simply results in a long-running connection.
                Disconnect(HangingRequestDisconnectReason.UserInitiated, null);
            }
            catch (XmlException ex)
            {
                // Thrown if server returned no XML document.
                Disconnect(HangingRequestDisconnectReason.UserInitiated, ex);
            }
            finally
            {
                if (responseCopy != null)
                {
                    await responseCopy.DisposeAsync().ConfigureAwait(false);
                }
            }
        }
        catch (ServiceLocalException exception)
        {
            Disconnect(HangingRequestDisconnectReason.Exception, exception);
        }
        catch (Exception exception)
        {
            Disconnect(HangingRequestDisconnectReason.Exception, exception);
        }
    }

    /// <summary>
    ///     Disconnects the request.
    /// </summary>
    internal void Disconnect()
    {
        lock (_lockObject)
        {
            Disconnect(HangingRequestDisconnectReason.UserInitiated, null);
        }
    }

    /// <summary>
    ///     Disconnects the request with the specified reason and exception.
    /// </summary>
    /// <param name="reason">The reason.</param>
    /// <param name="exception">The exception.</param>
    internal async void Disconnect(HangingRequestDisconnectReason reason, Exception? exception)
    {
        if (IsConnected)
        {
            _tokenSource.Cancel();
            // We do not care about exceptions here, as the ParseResponse handler provides a catch-all handler
            await _readTask;

            _response.Close();
            InternalOnDisconnect(reason, exception);
        }
    }


    /// <summary>
    ///     Perform any bookkeeping needed when we disconnect (cleanly or forcefully)
    /// </summary>
    /// <param name="reason"></param>
    /// <param name="exception"></param>
    private void InternalOnDisconnect(HangingRequestDisconnectReason reason, Exception? exception)
    {
        if (IsConnected)
        {
            IsConnected = false;

            OnDisconnect(this, new HangingRequestDisconnectEventArgs(reason, exception));
        }
    }

    /// <summary>
    ///     Reads any preamble data not part of the core response.
    /// </summary>
    /// <param name="ewsXmlReader">The EwsServiceXmlReader.</param>
    protected override void ReadPreamble(EwsServiceXmlReader ewsXmlReader)
    {
        // Do nothing.
    }

    /// <summary>
    ///     Reads any preamble data not part of the core response.
    /// </summary>
    /// <param name="ewsXmlReader">The EwsServiceXmlReader.</param>
    protected override System.Threading.Tasks.Task ReadPreambleAsync(EwsServiceXmlReader ewsXmlReader)
    {
        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <summary>
    ///     Callback delegate to handle asynchronous responses.
    /// </summary>
    /// <param name="response">Response received from the server</param>
    internal delegate void HandleResponseObject(object response);

    /// <summary>
    ///     Delegate method to handle a hanging request disconnection.
    /// </summary>
    /// <param name="sender">The object invoking the delegate.</param>
    /// <param name="args">Event data.</param>
    internal delegate void HangingRequestDisconnectHandler(object sender, HangingRequestDisconnectEventArgs args);
}
