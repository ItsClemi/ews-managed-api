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

using System.Globalization;
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

internal sealed class SecurityTimestamp
{
    //  Pulled from SecurityProtocolFactory
    //
    internal const string DefaultTimestampValidityDurationString = "00:05:00";

    internal const string DefaultFormat = "yyyy-MM-ddTHH:mm:ss.fffZ";

    internal static readonly TimeSpan DefaultTimestampValidityDuration =
        TimeSpan.Parse(DefaultTimestampValidityDurationString);

    //                            012345678901234567890123
    internal static readonly TimeSpan DefaultTimeToLive = DefaultTimestampValidityDuration;
    private readonly byte[] _digest;
    private readonly string _digestAlgorithm;
    private readonly string _id;
    private char[]? _computedCreationTimeUtc;
    private char[]? _computedExpiryTimeUtc;
    private DateTime _creationTimeUtc;
    private DateTime _expiryTimeUtc;

    public DateTime CreationTimeUtc => _creationTimeUtc;

    public DateTime ExpiryTimeUtc => _expiryTimeUtc;

    public string Id => _id;

    public string DigestAlgorithm => _digestAlgorithm;

    internal SecurityTimestamp(
        DateTime creationTimeUtc,
        DateTime expiryTimeUtc,
        string id,
        string? digestAlgorithm = null,
        byte[]? digest = null
    )
    {
        EwsUtilities.Assert(
            creationTimeUtc.Kind == DateTimeKind.Utc,
            "SecurityTimestamp.ctor",
            "creation time must be in UTC"
        );
        EwsUtilities.Assert(
            expiryTimeUtc.Kind == DateTimeKind.Utc,
            "SecurityTimestamp.ctor",
            "expiry time must be in UTC"
        );

        if (creationTimeUtc > expiryTimeUtc)
        {
            throw new ArgumentOutOfRangeException(nameof(expiryTimeUtc));
        }

        _creationTimeUtc = creationTimeUtc;
        _expiryTimeUtc = expiryTimeUtc;
        _id = id;

        _digestAlgorithm = digestAlgorithm;
        _digest = digest;
    }

    internal byte[] GetDigest()
    {
        return _digest;
    }

    internal char[] GetCreationTimeChars()
    {
        if (_computedCreationTimeUtc == null)
        {
            _computedCreationTimeUtc = ToChars(ref _creationTimeUtc);
        }

        return _computedCreationTimeUtc;
    }

    internal char[] GetExpiryTimeChars()
    {
        if (_computedExpiryTimeUtc == null)
        {
            _computedExpiryTimeUtc = ToChars(ref _expiryTimeUtc);
        }

        return _computedExpiryTimeUtc;
    }

    private static char[] ToChars(ref DateTime utcTime)
    {
        var buffer = new char[DefaultFormat.Length];
        var offset = 0;

        ToChars(utcTime.Year, buffer, ref offset, 4);
        buffer[offset++] = '-';

        ToChars(utcTime.Month, buffer, ref offset, 2);
        buffer[offset++] = '-';

        ToChars(utcTime.Day, buffer, ref offset, 2);
        buffer[offset++] = 'T';

        ToChars(utcTime.Hour, buffer, ref offset, 2);
        buffer[offset++] = ':';

        ToChars(utcTime.Minute, buffer, ref offset, 2);
        buffer[offset++] = ':';

        ToChars(utcTime.Second, buffer, ref offset, 2);
        buffer[offset++] = '.';

        ToChars(utcTime.Millisecond, buffer, ref offset, 3);
        buffer[offset++] = 'Z';

        return buffer;
    }

    private static void ToChars(int n, char[] buffer, ref int offset, int count)
    {
        for (var i = offset + count - 1; i >= offset; i--)
        {
            buffer[i] = (char)('0' + n % 10);
            n /= 10;
        }

        EwsUtilities.Assert(n == 0, "SecurityTimestamp.ToChars", "Overflow in encoding timestamp field");
        offset += count;
    }

    public override string ToString()
    {
        return string.Format(
            CultureInfo.InvariantCulture,
            "SecurityTimestamp: Id={0}, CreationTimeUtc={1}, ExpirationTimeUtc={2}",
            Id,
            XmlConvert.ToString(CreationTimeUtc, XmlDateTimeSerializationMode.RoundtripKind),
            XmlConvert.ToString(ExpiryTimeUtc, XmlDateTimeSerializationMode.RoundtripKind)
        );
    }
}
