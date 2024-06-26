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
///     Represents mailbox query object.
/// </summary>
[PublicAPI]
public sealed class MailboxQuery
{
    /// <summary>
    ///     Search query
    /// </summary>
    public string Query { get; set; }

    /// <summary>
    ///     Set of mailbox and scope pair
    /// </summary>
    public MailboxSearchScope[] MailboxSearchScopes { get; set; }

    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="query">Search query</param>
    /// <param name="searchScopes">Set of mailbox and scope pair</param>
    public MailboxQuery(string query, MailboxSearchScope[] searchScopes)
    {
        Query = query;
        MailboxSearchScopes = searchScopes;
    }
}

/// <summary>
///     Represents mailbox search scope object.
/// </summary>
[PublicAPI]
public sealed class MailboxSearchScope
{
    /// <summary>
    ///     Mailbox
    /// </summary>
    public string Mailbox { get; set; }

    /// <summary>
    ///     Search scope
    /// </summary>
    public MailboxSearchLocation SearchScope { get; set; }

    /// <summary>
    ///     Search scope type
    /// </summary>
    internal MailboxSearchScopeType SearchScopeType { get; set; } = MailboxSearchScopeType.LegacyExchangeDN;

    /// <summary>
    ///     Gets the extended data.
    /// </summary>
    /// <value>The extended data.</value>
    public ExtendedAttributes? ExtendedAttributes { get; private set; }

    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="mailbox">Mailbox</param>
    /// <param name="searchScope">Search scope</param>
    public MailboxSearchScope(string mailbox, MailboxSearchLocation searchScope = MailboxSearchLocation.All)
    {
        Mailbox = mailbox;
        SearchScope = searchScope;
        ExtendedAttributes = new ExtendedAttributes();
    }
}

/// <summary>
///     Represents mailbox object for preview item.
/// </summary>
[PublicAPI]
public sealed class PreviewItemMailbox
{
    /// <summary>
    ///     Mailbox id
    /// </summary>
    public string MailboxId { get; set; }

    /// <summary>
    ///     Primary smtp address
    /// </summary>
    public string PrimarySmtpAddress { get; set; }

    /// <summary>
    ///     Constructor
    /// </summary>
    public PreviewItemMailbox()
    {
    }

    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="mailboxId">Mailbox id</param>
    /// <param name="primarySmtpAddress">Primary smtp address</param>
    public PreviewItemMailbox(string mailboxId, string primarySmtpAddress)
    {
        MailboxId = mailboxId;
        PrimarySmtpAddress = primarySmtpAddress;
    }
}
