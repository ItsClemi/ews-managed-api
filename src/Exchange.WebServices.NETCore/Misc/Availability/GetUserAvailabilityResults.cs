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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the results of a GetUserAvailability operation.
/// </summary>
[PublicAPI]
public sealed class GetUserAvailabilityResults
{
    /// <summary>
    ///     Gets or sets the suggestions response for the requested meeting time.
    /// </summary>
    internal SuggestionsResponse? SuggestionsResponse { get; set; }

    /// <summary>
    ///     Gets a collection of AttendeeAvailability objects representing availability information for each of the specified
    ///     attendees.
    /// </summary>
    public ServiceResponseCollection<AttendeeAvailability> AttendeesAvailability { get; internal set; }

    /// <summary>
    ///     Gets a collection of suggested meeting times for the specified time period.
    /// </summary>
    public Collection<Suggestion>? Suggestions
    {
        get
        {
            if (SuggestionsResponse == null)
            {
                return null;
            }

            SuggestionsResponse.ThrowIfNecessary();

            return SuggestionsResponse.Suggestions;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetUserAvailabilityResults" /> class.
    /// </summary>
    internal GetUserAvailabilityResults()
    {
    }
}
