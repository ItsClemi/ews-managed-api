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
///     Represents an event that applies to a folder.
/// </summary>
[PublicAPI]
public class FolderEvent : NotificationEvent
{
    /// <summary>
    ///     Gets the Id of the folder this event applies to.
    /// </summary>
    public FolderId FolderId { get; private set; }

    /// <summary>
    ///     Gets the Id of the folder that was moved or copied. OldFolderId is only meaningful
    ///     when EventType is equal to either EventType.Moved or EventType.Copied. For all
    ///     other event types, OldFolderId is null.
    /// </summary>
    public FolderId OldFolderId { get; private set; }

    /// <summary>
    ///     Gets the new number of unread messages. This is is only meaningful when
    ///     EventType is equal to EventType.Modified. For all other event types,
    ///     UnreadCount is null.
    /// </summary>
    public int? UnreadCount { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderEvent" /> class.
    /// </summary>
    /// <param name="eventType">Type of the event.</param>
    /// <param name="timestamp">The event timestamp.</param>
    internal FolderEvent(EventType eventType, DateTime timestamp)
        : base(eventType, timestamp)
    {
    }

    /// <summary>
    ///     Load from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void InternalLoadFromXml(EwsServiceXmlReader reader)
    {
        base.InternalLoadFromXml(reader);

        FolderId = new FolderId();
        FolderId.LoadFromXml(reader, reader.LocalName);

        reader.Read();

        ParentFolderId = new FolderId();
        ParentFolderId.LoadFromXml(reader, XmlElementNames.ParentFolderId);

        switch (EventType)
        {
            case EventType.Moved:
            case EventType.Copied:
            {
                reader.Read();

                OldFolderId = new FolderId();
                OldFolderId.LoadFromXml(reader, reader.LocalName);

                reader.Read();

                OldParentFolderId = new FolderId();
                OldParentFolderId.LoadFromXml(reader, reader.LocalName);
                break;
            }
            case EventType.Modified:
            {
                reader.Read();
                if (reader.IsStartElement())
                {
                    reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.UnreadCount);
                    UnreadCount = int.Parse(reader.ReadValue());
                }

                break;
            }
        }
    }
}
