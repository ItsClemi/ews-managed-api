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
///     Represents a connection to an ongoing stream of events.
/// </summary>
[PublicAPI]
public sealed class StreamingSubscriptionConnection : IDisposable
{
    /// <summary>
    ///     Represents a delegate that is invoked when notifications are received from the server
    /// </summary>
    /// <param name="sender">The StreamingSubscriptionConnection instance that received the events.</param>
    /// <param name="args">The event data.</param>
    public delegate void NotificationEventDelegate(object sender, NotificationEventArgs args);

    /// <summary>
    ///     Represents a delegate that is invoked when an error occurs within a streaming subscription connection.
    /// </summary>
    /// <param name="sender">The StreamingSubscriptionConnection instance within which the error occurred.</param>
    /// <param name="args">The event data.</param>
    public delegate void SubscriptionErrorDelegate(object sender, SubscriptionErrorEventArgs args);

    /// <summary>
    ///     connection lifetime, in minutes
    /// </summary>
    private readonly int _connectionTimeout;

    /// <summary>
    ///     Lock object
    /// </summary>
    private readonly object _lockObject = new();

    /// <summary>
    ///     Currently used instance of a GetStreamingEventsRequest connected to EWS.
    /// </summary>
    private GetStreamingEventsRequest? _currentHangingRequest;

    /// <summary>
    ///     Value indicating whether the class is disposed.
    /// </summary>
    private bool _isDisposed;

    /// <summary>
    ///     ExchangeService instance used to make the EWS call.
    /// </summary>
    private ExchangeService _session;

    /// <summary>
    ///     Mapping of streaming id to subscriptions currently on the connection.
    /// </summary>
    private Dictionary<string, StreamingSubscription> _subscriptions = new();

    /// <summary>
    ///     Getting the current subscriptions in this connection.
    /// </summary>
    public IEnumerable<StreamingSubscription> CurrentSubscriptions
    {
        get
        {
            var result = new List<StreamingSubscription>();
            lock (_lockObject)
            {
                result.AddRange(_subscriptions.Values);
            }

            return result;
        }
    }

    /// <summary>
    ///     Gets a value indicating whether this connection is opened
    /// </summary>
    public bool IsOpen
    {
        get
        {
            ThrowIfDisposed();
            if (_currentHangingRequest == null)
            {
                return false;
            }

            return _currentHangingRequest.IsConnected;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="StreamingSubscriptionConnection" /> class.
    /// </summary>
    /// <param name="service">The ExchangeService instance this connection uses to connect to the server.</param>
    /// <param name="lifetime">
    ///     The maximum time, in minutes, the connection will remain open. Lifetime must be between 1 and
    ///     30.
    /// </param>
    public StreamingSubscriptionConnection(ExchangeService service, int lifetime)
    {
        EwsUtilities.ValidateParam(service);

        EwsUtilities.ValidateClassVersion(service, ExchangeVersion.Exchange2010_SP1, GetType().Name);

        if (lifetime < 1 || lifetime > 30)
        {
            throw new ArgumentOutOfRangeException(nameof(lifetime));
        }

        _session = service;
        _connectionTimeout = lifetime;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="StreamingSubscriptionConnection" /> class.
    /// </summary>
    /// <param name="service">The ExchangeService instance this connection uses to connect to the server.</param>
    /// <param name="subscriptions">The streaming subscriptions this connection is receiving events for.</param>
    /// <param name="lifetime">
    ///     The maximum time, in minutes, the connection will remain open. Lifetime must be between 1 and
    ///     30.
    /// </param>
    public StreamingSubscriptionConnection(
        ExchangeService service,
        IEnumerable<StreamingSubscription> subscriptions,
        int lifetime
    )
        : this(service, lifetime)
    {
        EwsUtilities.ValidateParamCollection(subscriptions);

        foreach (var subscription in subscriptions)
        {
            _subscriptions.Add(subscription.Id, subscription);
        }
    }

    /// <summary>
    ///     Occurs when notifications are received from the server.
    /// </summary>
    public event NotificationEventDelegate? OnNotificationEvent;

    /// <summary>
    ///     Occurs when a subscription encounters an error.
    /// </summary>
    public event SubscriptionErrorDelegate? OnSubscriptionError;

    /// <summary>
    ///     Occurs when a streaming subscription connection is disconnected from the server.
    /// </summary>
    public event SubscriptionErrorDelegate? OnDisconnect;

    /// <summary>
    ///     Adds a subscription to this connection.
    /// </summary>
    /// <param name="subscription">The subscription to add.</param>
    /// <exception cref="InvalidOperationException">Thrown when AddSubscription is called while connected.</exception>
    public void AddSubscription(StreamingSubscription subscription)
    {
        ThrowIfDisposed();

        EwsUtilities.ValidateParam(subscription);

        ValidateConnectionState(false, Strings.CannotAddSubscriptionToLiveConnection);

        lock (_lockObject)
        {
            _subscriptions.TryAdd(subscription.Id, subscription);
        }
    }

    /// <summary>
    ///     Removes the specified streaming subscription from the connection.
    /// </summary>
    /// <param name="subscription">The subscription to remove.</param>
    /// <exception cref="InvalidOperationException">Thrown when RemoveSubscription is called while connected.</exception>
    public void RemoveSubscription(StreamingSubscription subscription)
    {
        ThrowIfDisposed();

        EwsUtilities.ValidateParam(subscription);

        ValidateConnectionState(false, Strings.CannotRemoveSubscriptionFromLiveConnection);

        lock (_lockObject)
        {
            _subscriptions.Remove(subscription.Id);
        }
    }

    /// <summary>
    ///     Opens this connection so it starts receiving events from the server.
    ///     This results in a long-standing call to EWS.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when Open is called while connected.</exception>
    public void Open()
    {
        lock (_lockObject)
        {
            ThrowIfDisposed();

            ValidateConnectionState(false, Strings.CannotCallConnectDuringLiveConnection);

            if (_subscriptions.Count == 0)
            {
                throw new ServiceLocalException(Strings.NoSubscriptionsOnConnection);
            }

            _currentHangingRequest = new GetStreamingEventsRequest(
                _session,
                HandleServiceResponseObject,
                _subscriptions.Keys,
                _connectionTimeout
            );

            _currentHangingRequest.OnDisconnect += OnRequestDisconnect;

            _currentHangingRequest.InternalExecute();
        }
    }

    /// <summary>
    ///     Called when the request is disconnected.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="args">
    ///     The <see cref="Microsoft.Exchange.WebServices.Data.HangingRequestDisconnectEventArgs" /> instance
    ///     containing the event data.
    /// </param>
    private void OnRequestDisconnect(object sender, HangingRequestDisconnectEventArgs args)
    {
        InternalOnDisconnect(args.Exception);
    }

    /// <summary>
    ///     Closes this connection so it stops receiving events from the server.
    ///     This terminates a long-standing call to EWS.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when Close is called while not connected.</exception>
    public void Close()
    {
        lock (_lockObject)
        {
            ThrowIfDisposed();

            ValidateConnectionState(true, Strings.CannotCallDisconnectWithNoLiveConnection);

            // Further down in the stack, this will result in a call to our OnRequestDisconnect event handler,
            // doing the necessary cleanup.
            _currentHangingRequest?.Disconnect();
        }
    }

    /// <summary>
    ///     Internal helper method called when the request disconnects.
    /// </summary>
    /// <param name="ex">The exception that caused the disconnection. May be null.</param>
    private void InternalOnDisconnect(Exception? ex)
    {
        _currentHangingRequest = null;

        OnDisconnect?.Invoke(this, new SubscriptionErrorEventArgs(null, ex));
    }

    /// <summary>
    ///     Validates the state of the connection.
    /// </summary>
    /// <param name="isConnectedExpected">Value indicating whether we expect to be currently connected.</param>
    /// <param name="errorMessage">The error message.</param>
    private void ValidateConnectionState(bool isConnectedExpected, string errorMessage)
    {
        if ((isConnectedExpected && !IsOpen) || (!isConnectedExpected && IsOpen))
        {
            throw new ServiceLocalException(errorMessage);
        }
    }

    /// <summary>
    ///     Handles the service response object.
    /// </summary>
    /// <param name="response">The response.</param>
    private void HandleServiceResponseObject(object response)
    {
        if (response is not GetStreamingEventsResponse gseResponse)
        {
            throw new ArgumentException(null, nameof(response));
        }

        if (gseResponse.Result == ServiceResult.Success || gseResponse.Result == ServiceResult.Warning)
        {
            if (gseResponse.Results.Notifications.Count > 0)
            {
                // We got notifications; dole them out.
                IssueNotificationEvents(gseResponse);
            }
            // This was just a heartbeat, nothing to do here.
        }
        else if (gseResponse.Result == ServiceResult.Error)
        {
            if (gseResponse.ErrorSubscriptionIds == null || gseResponse.ErrorSubscriptionIds.Count == 0)
            {
                // General error
                IssueGeneralFailure(gseResponse);
            }
            else
            {
                // subscription-specific errors
                IssueSubscriptionFailures(gseResponse);
            }
        }
    }

    /// <summary>
    ///     Issues the subscription failures.
    /// </summary>
    /// <param name="gseResponse">The GetStreamingEvents response.</param>
    private void IssueSubscriptionFailures(GetStreamingEventsResponse gseResponse)
    {
        var exception = new ServiceResponseException(gseResponse);

        foreach (var id in gseResponse.ErrorSubscriptionIds)
        {
            StreamingSubscription? subscription = null;

            lock (_lockObject)
            {
                // Client can do any good or bad things in the below event handler
                _subscriptions?.TryGetValue(id, out subscription);
            }

            if (subscription != null)
            {
                var eventArgs = new SubscriptionErrorEventArgs(subscription, exception);

                OnSubscriptionError?.Invoke(this, eventArgs);
            }

            if (gseResponse.ErrorCode != ServiceError.ErrorMissedNotificationEvents)
            {
                // Client can do any good or bad things in the above event handler
                lock (_lockObject)
                {
                    // We are no longer servicing the subscription.
                    _subscriptions?.Remove(id);
                }
            }
        }
    }

    /// <summary>
    ///     Issues the general failure.
    /// </summary>
    /// <param name="gseResponse">The GetStreamingEvents response.</param>
    private void IssueGeneralFailure(GetStreamingEventsResponse gseResponse)
    {
        var eventArgs = new SubscriptionErrorEventArgs(null, new ServiceResponseException(gseResponse));

        OnSubscriptionError?.Invoke(this, eventArgs);
    }

    /// <summary>
    ///     Issues the notification events.
    /// </summary>
    /// <param name="gseResponse">The GetStreamingEvents response.</param>
    private void IssueNotificationEvents(GetStreamingEventsResponse gseResponse)
    {
        foreach (var events in gseResponse.Results.Notifications)
        {
            StreamingSubscription? subscription = null;

            lock (_lockObject)
            {
                // Client can do any good or bad things in the below event handler
                _subscriptions?.TryGetValue(events.SubscriptionId, out subscription);
            }

            if (subscription != null)
            {
                var eventArgs = new NotificationEventArgs(subscription, events.Events);

                OnNotificationEvent?.Invoke(this, eventArgs);
            }
        }
    }


    #region IDisposable Members

    /// <summary>
    ///     Finalizes an instance of the StreamingSubscriptionConnection class.
    /// </summary>
    ~StreamingSubscriptionConnection()
    {
        Dispose(false);
    }

    /// <summary>
    ///     Frees resources associated with this StreamingSubscriptionConnection.
    /// </summary>
    public void Dispose()
    {
        GC.SuppressFinalize(this);

        Dispose(true);
    }

    /// <summary>
    ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    /// <param name="suppressFinalizer">Value indicating whether to suppress the garbage collector's finalizer..</param>
    private void Dispose(bool suppressFinalizer)
    {
        lock (_lockObject)
        {
            if (!_isDisposed)
            {
                if (_currentHangingRequest != null && _currentHangingRequest.IsConnected)
                {
                    _currentHangingRequest.Disconnect();
                }

                _currentHangingRequest = null;
                _subscriptions = null;
                _session = null;

                _isDisposed = true;
            }
        }
    }

    /// <summary>
    ///     Throws if disposed.
    /// </summary>
    private void ThrowIfDisposed()
    {
        if (_isDisposed)
        {
            throw new ObjectDisposedException(GetType().FullName);
        }
    }

    #endregion
}
