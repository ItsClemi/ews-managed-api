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
///     Represents the options of a GetAvailability request.
/// </summary>
[PublicAPI]
public sealed class AvailabilityOptions
{
    private int _goodSuggestionThreshold = 25;
    private int _maximumNonWorkHoursSuggestionsPerDay;
    private int _maximumSuggestionsPerDay = 10;
    private int _meetingDuration = 60;
    private int _mergedFreeBusyInterval = 30;

    /// <summary>
    ///     Gets or sets the time difference between two successive slots in a FreeBusyMerged view.
    ///     MergedFreeBusyInterval must be between 5 and 1440. The default value is 30.
    /// </summary>
    public int MergedFreeBusyInterval
    {
        get => _mergedFreeBusyInterval;

        set
        {
            if (value < 5 || value > 1440)
            {
                throw new ArgumentException(
                    string.Format(Strings.InvalidPropertyValueNotInRange, "MergedFreeBusyInterval", 5, 1440),
                    nameof(value)
                );
            }

            _mergedFreeBusyInterval = value;
        }
    }

    /// <summary>
    ///     Gets or sets the requested type of free/busy view. The default value is FreeBusyViewType.Detailed.
    /// </summary>
    public FreeBusyViewType RequestedFreeBusyView { get; set; } = FreeBusyViewType.Detailed;

    /// <summary>
    ///     Gets or sets the percentage of attendees that must have the time period open for the time period to qualify as a
    ///     good suggested meeting time.
    ///     GoodSuggestionThreshold must be between 1 and 49. The default value is 25.
    /// </summary>
    public int GoodSuggestionThreshold
    {
        get => _goodSuggestionThreshold;

        set
        {
            if (value < 1 || value > 49)
            {
                throw new ArgumentException(
                    string.Format(Strings.InvalidPropertyValueNotInRange, "GoodSuggestionThreshold", 1, 49),
                    nameof(value)
                );
            }

            _goodSuggestionThreshold = value;
        }
    }

    /// <summary>
    ///     Gets or sets the number of suggested meeting times that should be returned per day.
    ///     MaximumSuggestionsPerDay must be between 0 and 48. The default value is 10.
    /// </summary>
    public int MaximumSuggestionsPerDay
    {
        get => _maximumSuggestionsPerDay;

        set
        {
            if (value < 0 || value > 48)
            {
                throw new ArgumentException(
                    string.Format(Strings.InvalidPropertyValueNotInRange, "MaximumSuggestionsPerDay", 0, 48),
                    nameof(value)
                );
            }

            _maximumSuggestionsPerDay = value;
        }
    }

    /// <summary>
    ///     Gets or sets the number of suggested meeting times outside regular working hours per day.
    ///     MaximumNonWorkHoursSuggestionsPerDay must be between 0 and 48. The default value is 0.
    /// </summary>
    public int MaximumNonWorkHoursSuggestionsPerDay
    {
        get => _maximumNonWorkHoursSuggestionsPerDay;

        set
        {
            if (value < 0 || value > 48)
            {
                throw new ArgumentException(
                    string.Format(
                        Strings.InvalidPropertyValueNotInRange,
                        "MaximumNonWorkHoursSuggestionsPerDay",
                        0,
                        48
                    ),
                    nameof(value)
                );
            }

            _maximumNonWorkHoursSuggestionsPerDay = value;
        }
    }

    /// <summary>
    ///     Gets or sets the duration, in minutes, of the meeting for which to obtain suggestions.
    ///     MeetingDuration must be between 30 and 1440. The default value is 60.
    /// </summary>
    public int MeetingDuration
    {
        get => _meetingDuration;

        set
        {
            if (value < 30 || value > 1440)
            {
                throw new ArgumentException(
                    string.Format(Strings.InvalidPropertyValueNotInRange, "MeetingDuration", 30, 1440),
                    nameof(value)
                );
            }

            _meetingDuration = value;
        }
    }

    /// <summary>
    ///     Gets or sets the minimum quality of suggestions that should be returned.
    ///     The default is SuggestionQuality.Fair.
    /// </summary>
    public SuggestionQuality MinimumSuggestionQuality { get; set; } = SuggestionQuality.Fair;

    /// <summary>
    ///     Gets or sets the time window for which detailed information about suggested meeting times should be returned.
    /// </summary>
    public TimeWindow? DetailedSuggestionsWindow { get; set; }

    /// <summary>
    ///     Gets or sets the start time of a meeting that you want to update with the suggested meeting times.
    /// </summary>
    public DateTime? CurrentMeetingTime { get; set; }

    /// <summary>
    ///     Gets or sets the global object Id of a meeting that will be modified based on the data returned by
    ///     GetUserAvailability.
    /// </summary>
    public string GlobalObjectId { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AvailabilityOptions" /> class.
    /// </summary>
    public AvailabilityOptions()
    {
    }

    /// <summary>
    ///     Validates this instance against the specified time window.
    /// </summary>
    /// <param name="timeWindow">The time window.</param>
    internal void Validate(TimeSpan timeWindow)
    {
        if (TimeSpan.FromMinutes(MergedFreeBusyInterval) > timeWindow)
        {
            throw new ArgumentException(
                Strings.MergedFreeBusyIntervalMustBeSmallerThanTimeWindow,
                nameof(MergedFreeBusyInterval)
            );
        }

        EwsUtilities.ValidateParamAllowNull(DetailedSuggestionsWindow);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="request">The request being emitted.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, GetUserAvailabilityRequest request)
    {
        if (request.IsFreeBusyViewRequested)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FreeBusyViewOptions);

            request.TimeWindow.WriteToXmlUnscopedDatesOnly(writer, XmlElementNames.TimeWindow);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.MergedFreeBusyIntervalInMinutes,
                MergedFreeBusyInterval
            );

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RequestedView, RequestedFreeBusyView);

            writer.WriteEndElement(); // FreeBusyViewOptions
        }

        if (request.IsSuggestionsViewRequested)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.SuggestionsViewOptions);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GoodThreshold, GoodSuggestionThreshold);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MaximumResultsByDay, MaximumSuggestionsPerDay);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.MaximumNonWorkHourResultsByDay,
                MaximumNonWorkHoursSuggestionsPerDay
            );

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MeetingDurationInMinutes, MeetingDuration);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.MinimumSuggestionQuality,
                MinimumSuggestionQuality
            );

            var timeWindowToSerialize = DetailedSuggestionsWindow ?? request.TimeWindow;

            timeWindowToSerialize.WriteToXmlUnscopedDatesOnly(writer, XmlElementNames.DetailedSuggestionsWindow);

            if (CurrentMeetingTime.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.CurrentMeetingTime,
                    CurrentMeetingTime.Value
                );
            }

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GlobalObjectId, GlobalObjectId);

            writer.WriteEndElement(); // SuggestionsViewOptions
        }
    }
}
