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
///     Lists different keys that can be passed to the people query context dictionary
/// </summary>
[PublicAPI]
public sealed class PeopleQueryContextKeys
{
    /// <summary> Topic Query </summary>
    public const string TopicQuery = "TopicQuery";

    /// <summary> Guid for client session </summary>
    public const string ClientSessionId = "ClientSessionId";

    /// <summary> Client Flight Number </summary>
    public const string ClientFlightNumber = "ClientFlightNumber";

    /// <summary> User Agent </summary>
    public const string UserAgent = "UserAgent";

    /// <summary> App Name </summary>
    public const string AppName = "AppName";

    /// <summary> App Scenario </summary>
    public const string AppScenario = "AppScenario";
}
