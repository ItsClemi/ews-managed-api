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

using System.ComponentModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the base response class for synchronuization operations.
/// </summary>
/// <typeparam name="TServiceObject">ServiceObject type.</typeparam>
/// <typeparam name="TChange">Change type.</typeparam>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class SyncResponse<TServiceObject, TChange> : ServiceResponse
    where TServiceObject : ServiceObject
    where TChange : Change
{
    private readonly PropertySet _propertySet;

    /// <summary>
    ///     Gets a list of changes that occurred on the synchronized folder.
    /// </summary>
    public ChangeCollection<TChange> Changes { get; } = new();

    /// <summary>
    ///     Gets a value indicating whether this request returns full or summary properties.
    /// </summary>
    internal abstract bool SummaryPropertiesOnly { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SyncResponse&lt;TServiceObject, TChange&gt;" /> class.
    /// </summary>
    /// <param name="propertySet">Property set.</param>
    internal SyncResponse(PropertySet propertySet)
    {
        _propertySet = propertySet;

        EwsUtilities.Assert(_propertySet != null, "SyncResponse.ctor", "PropertySet should not be null");
    }

    /// <summary>
    ///     Gets the name of the includes last in range XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal abstract string GetIncludesLastInRangeXmlElementName();

    /// <summary>
    ///     Creates the change instance.
    /// </summary>
    /// <returns>TChange instance</returns>
    internal abstract TChange CreateChangeInstance();

    /// <summary>
    ///     Gets the name of the change element.
    /// </summary>
    /// <returns>Change element name.</returns>
    internal abstract string GetChangeElementName();

    /// <summary>
    ///     Gets the name of the change id element.
    /// </summary>
    /// <returns>Change id element name.</returns>
    internal abstract string GetChangeIdElementName();

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        Changes.SyncState = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.SyncState);
        Changes.MoreChangesAvailable = !reader.ReadElementValue<bool>(
            XmlNamespace.Messages,
            GetIncludesLastInRangeXmlElementName()
        );

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Changes);
        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    var change = CreateChangeInstance();

                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Create:
                        {
                            change.ChangeType = ChangeType.Create;
                            break;
                        }
                        case XmlElementNames.Update:
                        {
                            change.ChangeType = ChangeType.Update;
                            break;
                        }
                        case XmlElementNames.Delete:
                        {
                            change.ChangeType = ChangeType.Delete;
                            break;
                        }
                        case XmlElementNames.ReadFlagChange:
                        {
                            change.ChangeType = ChangeType.ReadFlagChange;
                            break;
                        }
                        default:
                        {
                            reader.SkipCurrentElement();
                            break;
                        }
                    }

                    if (change != null)
                    {
                        reader.Read();
                        reader.EnsureCurrentNodeIsStartElement();

                        switch (change.ChangeType)
                        {
                            case ChangeType.Delete:
                            case ChangeType.ReadFlagChange:
                            {
                                change.Id = change.CreateId();
                                change.Id.LoadFromXml(reader, change.Id.GetXmlElementName());

                                if (change.ChangeType == ChangeType.ReadFlagChange)
                                {
                                    reader.Read();
                                    reader.EnsureCurrentNodeIsStartElement();

                                    var itemChange = change as ItemChange;

                                    EwsUtilities.Assert(
                                        itemChange != null,
                                        "SyncResponse.ReadElementsFromXml",
                                        "ReadFlagChange is only valid on ItemChange"
                                    );

                                    itemChange.IsRead = reader.ReadElementValue<bool>(
                                        XmlNamespace.Types,
                                        XmlElementNames.IsRead
                                    );
                                }

                                break;
                            }
                            default:
                            {
                                change.ServiceObject = EwsUtilities.CreateEwsObjectFromXmlElementName<TServiceObject>(
                                    reader.Service,
                                    reader.LocalName
                                );

                                change.ServiceObject.LoadFromXml(reader, true, _propertySet, SummaryPropertiesOnly);
                                break;
                            }
                        }

                        reader.ReadEndElementIfNecessary(XmlNamespace.Types, change.ChangeType.ToString());

                        Changes.Add(change);
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Changes));
        }
    }
}
