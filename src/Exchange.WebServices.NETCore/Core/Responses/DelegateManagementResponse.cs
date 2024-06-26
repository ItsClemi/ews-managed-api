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
///     Represents the response to a delegate managent-related operation.
/// </summary>
internal class DelegateManagementResponse : ServiceResponse
{
    private readonly List<DelegateUser>? _delegateUsers;
    private readonly bool _readDelegateUsers;

    /// <summary>
    ///     Gets a collection of responses for each of the delegate users concerned by the operation.
    /// </summary>
    internal Collection<DelegateUserResponse> DelegateUserResponses { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegateManagementResponse" /> class.
    /// </summary>
    /// <param name="readDelegateUsers">if set to <c>true</c> [read delegate users].</param>
    /// <param name="delegateUsers">List of existing delegate users to load.</param>
    internal DelegateManagementResponse(bool readDelegateUsers, List<DelegateUser>? delegateUsers)
    {
        _readDelegateUsers = readDelegateUsers;
        _delegateUsers = delegateUsers;
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        if (ErrorCode == ServiceError.NoError)
        {
            DelegateUserResponses = new Collection<DelegateUserResponse>();

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages))
            {
                var delegateUserIndex = 0;
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.DelegateUserResponseMessageType))
                    {
                        DelegateUser? delegateUser = null;
                        if (_readDelegateUsers && _delegateUsers != null)
                        {
                            delegateUser = _delegateUsers[delegateUserIndex];
                        }

                        var delegateUserResponse = new DelegateUserResponse(_readDelegateUsers, delegateUser);

                        delegateUserResponse.LoadFromXml(reader, XmlElementNames.DelegateUserResponseMessageType);

                        DelegateUserResponses.Add(delegateUserResponse);

                        delegateUserIndex++;
                    }
                } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages));
            }
        }
    }
}
