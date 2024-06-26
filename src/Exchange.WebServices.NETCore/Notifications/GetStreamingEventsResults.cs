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

using System.Collections.ObjectModel;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of notification events.
/// </summary>
internal sealed class GetStreamingEventsResults
{
    /// <summary>
    ///     Gets the notification collection.
    /// </summary>
    /// <value>The notification collection.</value>
    internal Collection<NotificationGroup> Notifications { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetStreamingEventsResults" /> class.
    /// </summary>
    internal GetStreamingEventsResults()
    {
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Notification);

        do
        {
            var notifications = new NotificationGroup
            {
                SubscriptionId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SubscriptionId),
                Events = new Collection<NotificationEvent>(),
            };

            lock (this)
            {
                Notifications.Add(notifications);
            }

            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    var eventElementName = reader.LocalName;

                    if (GetEventsResults.XmlElementNameToEventTypeMap.TryGetValue(eventElementName, out var eventType))
                    {
                        if (eventType == EventType.Status)
                        {
                            // We don't need to return status events
                            reader.ReadEndElementIfNecessary(XmlNamespace.Types, eventElementName);
                        }
                        else
                        {
                            LoadNotificationEventFromXml(reader, eventElementName, eventType, notifications);
                        }
                    }
                    else
                    {
                        reader.SkipCurrentElement();
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notification));

            reader.Read();
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notifications));
    }

    /// <summary>
    ///     Loads a notification event from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="eventElementName">Name of the event XML element.</param>
    /// <param name="eventType">Type of the event.</param>
    /// <param name="notifications">Collection of notifications</param>
    private static void LoadNotificationEventFromXml(
        EwsServiceXmlReader reader,
        string eventElementName,
        EventType eventType,
        NotificationGroup notifications
    )
    {
        var timestamp = reader.ReadElementValue<DateTime>(XmlNamespace.Types, XmlElementNames.TimeStamp);

        NotificationEvent notificationEvent;

        reader.Read();

        if (reader.LocalName == XmlElementNames.FolderId)
        {
            notificationEvent = new FolderEvent(eventType, timestamp);
        }
        else
        {
            notificationEvent = new ItemEvent(eventType, timestamp);
        }

        notificationEvent.LoadFromXml(reader, eventElementName);
        notifications.Events.Add(notificationEvent);
    }

    /// <summary>
    ///     Structure to track a subscription and its associated notification events.
    /// </summary>
    internal struct NotificationGroup
    {
        /// <summary>
        ///     Subscription Id
        /// </summary>
        internal string SubscriptionId;

        /// <summary>
        ///     Events in the response associated with the subscription id.
        /// </summary>
        internal Collection<NotificationEvent> Events;
    }
}
