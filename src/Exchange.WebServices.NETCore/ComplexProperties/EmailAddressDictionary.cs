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
using System.Diagnostics.CodeAnalysis;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a dictionary of e-mail addresses.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class EmailAddressDictionary : DictionaryProperty<EmailAddressKey, EmailAddressEntry>
{
    /// <summary>
    ///     Gets or sets the e-mail address at the specified key.
    /// </summary>
    /// <param name="key">The key of the e-mail address to get or set.</param>
    /// <returns>The e-mail address at the specified key.</returns>
    public EmailAddress? this[EmailAddressKey key]
    {
        get => Entries[key].EmailAddress;

        set
        {
            if (value == null)
            {
                InternalRemove(key);
            }
            else
            {
                if (Entries.TryGetValue(key, out var entry))
                {
                    entry.EmailAddress = value;
                    Changed();
                }
                else
                {
                    entry = new EmailAddressEntry(key, value);
                    InternalAdd(entry);
                }
            }
        }
    }

    /// <summary>
    ///     Gets the field URI.
    /// </summary>
    /// <returns>Field URI.</returns>
    internal override string GetFieldUri()
    {
        return "contacts:EmailAddress";
    }

    /// <summary>
    ///     Creates instance of dictionary entry.
    /// </summary>
    /// <returns>New instance.</returns>
    internal override EmailAddressEntry CreateEntryInstance()
    {
        return new EmailAddressEntry();
    }

    /// <summary>
    ///     Tries to get the e-mail address associated with the specified key.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <param name="emailAddress">
    ///     When this method returns, contains the e-mail address associated with the specified key,
    ///     if the key is found; otherwise, null. This parameter is passed uninitialized.
    /// </param>
    /// <returns>
    ///     true if the Dictionary contains an e-mail address associated with the specified key; otherwise, false.
    /// </returns>
    public bool TryGetValue(EmailAddressKey key, [MaybeNullWhen(false)] out EmailAddress emailAddress)
    {
        if (Entries.TryGetValue(key, out var entry))
        {
            emailAddress = entry.EmailAddress;
            return true;
        }

        emailAddress = null;
        return false;
    }
}
