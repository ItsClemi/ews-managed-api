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
///     Represents an event as exposed by push and pull notifications.
/// </summary>
[PublicAPI]
public abstract class NotificationEvent
{
    /// <summary>
    ///     Gets the type of this event.
    /// </summary>
    public EventType EventType { get; }

    /// <summary>
    ///     Gets the date and time when the event occurred.
    /// </summary>
    public DateTime TimeStamp { get; }

    /// <summary>
    ///     Gets the Id of the parent folder of the item or folder this event applies to.
    /// </summary>
    public FolderId ParentFolderId { get; internal set; }

    /// <summary>
    ///     Gets the Id of the old parent folder of the item or folder this event applies to.
    ///     OldParentFolderId is only meaningful when EventType is equal to either EventType.Moved or
    ///     EventType.Copied. For all other event types, OldParentFolderId is null.
    /// </summary>
    public FolderId OldParentFolderId { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="NotificationEvent" /> class.
    /// </summary>
    /// <param name="eventType">Type of the event.</param>
    /// <param name="timestamp">The event timestamp.</param>
    internal NotificationEvent(EventType eventType, DateTime timestamp)
    {
        EventType = eventType;
        TimeStamp = timestamp;
    }

    /// <summary>
    ///     Load from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void InternalLoadFromXml(EwsServiceXmlReader reader)
    {
    }

    /// <summary>
    ///     Loads this NotificationEvent from XML.
    /// </summary>
    /// <param name="reader">The reader from which to read the notification event.</param>
    /// <param name="xmlElementName">The start XML element name of this notification event.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        InternalLoadFromXml(reader);

        reader.ReadEndElementIfNecessary(XmlNamespace.Types, xmlElementName);
    }
}
