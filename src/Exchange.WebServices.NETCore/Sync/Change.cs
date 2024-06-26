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
///     Represents a change as returned by a synchronization operation.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class Change
{
    /// <summary>
    ///     The Id of the service object the change applies to.
    /// </summary>
    private ServiceId _id;

    /// <summary>
    ///     Gets the type of the change.
    /// </summary>
    public ChangeType ChangeType { get; internal set; }

    /// <summary>
    ///     Gets or sets the service object the change applies to.
    /// </summary>
    internal ServiceObject? ServiceObject { get; set; }

    /// <summary>
    ///     Gets or sets the Id of the service object the change applies to.
    /// </summary>
    internal ServiceId Id
    {
        get => ServiceObject != null ? ServiceObject.GetId() : _id;
        set => _id = value;
    }

    /// <summary>
    ///     Initializes a new instance of Change.
    /// </summary>
    internal Change()
    {
    }

    /// <summary>
    ///     Creates an Id of the appropriate class.
    /// </summary>
    /// <returns>A ServiceId.</returns>
    internal abstract ServiceId CreateId();
}
