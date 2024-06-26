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
///     Represents a client token access request
/// </summary>
[PublicAPI]
public class ClientAccessTokenRequest : ComplexProperty
{
    /// <summary>
    ///     Gets the App Id.
    /// </summary>
    public string Id { get; }

    /// <summary>
    ///     Gets token type.
    /// </summary>
    public ClientAccessTokenType TokenType { get; }

    /// <summary>
    ///     Gets the token scope.
    /// </summary>
    public string? Scope { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ClientAccessTokenRequest" /> class.
    /// </summary>
    /// <param name="id">id</param>
    /// <param name="tokenType">The tokenType.</param>
    /// <param name="scope">The scope.</param>
    public ClientAccessTokenRequest(string id, ClientAccessTokenType tokenType, string? scope = null)
    {
        Id = id;
        TokenType = tokenType;
        Scope = scope;
    }
}
