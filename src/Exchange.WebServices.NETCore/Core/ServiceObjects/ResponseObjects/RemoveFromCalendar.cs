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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a response object created to remove a calendar item from a meeting cancellation.
/// </summary>
[ServiceObjectDefinition(XmlElementNames.RemoveItem, ReturnedByServer = false)]
internal sealed class RemoveFromCalendar : ServiceObject
{
    private readonly Item _referenceItem;

    /// <summary>
    ///     Initializes a new instance of the <see cref="RemoveFromCalendar" /> class.
    /// </summary>
    /// <param name="referenceItem">The reference item.</param>
    internal RemoveFromCalendar(Item referenceItem)
        : base(referenceItem.Service)
    {
        EwsUtilities.Assert(referenceItem != null, "RemoveFromCalendar.ctor", "referenceItem is null");

        referenceItem.ThrowIfThisIsNew();

        _referenceItem = referenceItem;
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return ResponseObjectSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Loads the specified set of properties on the object.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(
        PropertySet propertySet,
        CancellationToken token
    )
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     Deletes the object.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
    /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        CancellationToken token
    )
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     Create response object.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="token"></param>
    /// <returns>A list of items that were created or modified as a results of this operation.</returns>
    internal Task<List<Item>> InternalCreate(
        FolderId? parentFolderId,
        MessageDisposition? messageDisposition,
        CancellationToken token
    )
    {
        ((ItemId)PropertyBag[ResponseObjectSchema.ReferenceItemId]).Assign(_referenceItem.Id);

        return Service.InternalCreateResponseObject(this, parentFolderId, messageDisposition, token);
    }
}
