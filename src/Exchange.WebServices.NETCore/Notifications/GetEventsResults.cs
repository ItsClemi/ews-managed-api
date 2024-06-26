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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of notification events.
/// </summary>
[PublicAPI]
public sealed class GetEventsResults
{
    /// <summary>
    ///     Map XML element name to notification event type.
    /// </summary>
    /// <remarks>
    ///     If you add a new notification event type, you'll need to add a new entry to the dictionary here.
    /// </remarks>
    internal static readonly IReadOnlyDictionary<string, EventType> XmlElementNameToEventTypeMap =
        new Dictionary<string, EventType>
        {
            // @formatter:off
            { XmlElementNames.CopiedEvent, EventType.Copied },
            { XmlElementNames.CreatedEvent, EventType.Created },
            { XmlElementNames.DeletedEvent, EventType.Deleted },
            { XmlElementNames.ModifiedEvent, EventType.Modified },
            { XmlElementNames.MovedEvent, EventType.Moved },
            { XmlElementNames.NewMailEvent, EventType.NewMail },
            { XmlElementNames.StatusEvent, EventType.Status },
            { XmlElementNames.FreeBusyChangedEvent, EventType.FreeBusyChanged },
            // @formatter:on
        };

    /// <summary>
    ///     Gets the Id of the subscription the collection is associated with.
    /// </summary>
    internal string SubscriptionId { get; private set; }

    /// <summary>
    ///     Gets the subscription's previous watermark.
    /// </summary>
    internal string PreviousWatermark { get; private set; }

    /// <summary>
    ///     Gets the subscription's new watermark.
    /// </summary>
    internal string NewWatermark { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether more events are available on the Exchange server.
    /// </summary>
    internal bool MoreEventsAvailable { get; private set; }

    /// <summary>
    ///     Gets the collection of folder events.
    /// </summary>
    /// <value>The folder events.</value>
    public IEnumerable<FolderEvent> FolderEvents => AllEvents.OfType<FolderEvent>();

    /// <summary>
    ///     Gets the collection of item events.
    /// </summary>
    /// <value>The item events.</value>
    public IEnumerable<ItemEvent> ItemEvents => AllEvents.OfType<ItemEvent>();

    /// <summary>
    ///     Gets the collection of all events.
    /// </summary>
    /// <value>The events.</value>
    public Collection<NotificationEvent> AllEvents { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetEventsResults" /> class.
    /// </summary>
    internal GetEventsResults()
    {
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Notification);

        SubscriptionId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SubscriptionId);
        PreviousWatermark = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PreviousWatermark);
        MoreEventsAvailable = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.MoreEvents);

        do
        {
            reader.Read();

            if (reader.IsStartElement())
            {
                var eventElementName = reader.LocalName;

                if (XmlElementNameToEventTypeMap.TryGetValue(eventElementName, out var eventType))
                {
                    NewWatermark = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Watermark);

                    if (eventType == EventType.Status)
                    {
                        // We don't need to return status events
                        reader.ReadEndElementIfNecessary(XmlNamespace.Types, eventElementName);
                    }
                    else
                    {
                        LoadNotificationEventFromXml(reader, eventElementName, eventType);
                    }
                }
                else
                {
                    reader.SkipCurrentElement();
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notification));
    }

    /// <summary>
    ///     Loads a notification event from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="eventElementName">Name of the event XML element.</param>
    /// <param name="eventType">Type of the event.</param>
    private void LoadNotificationEventFromXml(EwsServiceXmlReader reader, string eventElementName, EventType eventType)
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
        AllEvents.Add(notificationEvent);
    }
}
