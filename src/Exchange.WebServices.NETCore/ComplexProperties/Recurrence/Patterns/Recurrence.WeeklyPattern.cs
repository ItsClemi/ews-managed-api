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

/// <content>
///     Contains nested type Recurrence.WeeklyPattern.
/// </content>
public abstract partial class Recurrence
{
    /// <summary>
    ///     Represents a recurrence pattern where each occurrence happens on specific days a specific number of weeks after the
    ///     previous one.
    /// </summary>
    [PublicAPI]
    public sealed class WeeklyPattern : IntervalPattern
    {
        private DayOfWeek? _firstDayOfWeek;

        /// <summary>
        ///     Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName => XmlElementNames.WeeklyRecurrence;

        /// <summary>
        ///     Gets the list of the days of the week when occurrences happen.
        /// </summary>
        public DayOfTheWeekCollection DaysOfTheWeek { get; } = new();

        /// <summary>
        ///     Gets or sets the first day of the week for this recurrence.
        /// </summary>
        public DayOfWeek FirstDayOfWeek
        {
            get => GetFieldValueOrThrowIfNull(_firstDayOfWeek, "FirstDayOfWeek");
            set => SetFieldValue(ref _firstDayOfWeek, value);
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="WeeklyPattern" /> class.
        /// </summary>
        public WeeklyPattern()
        {
            DaysOfTheWeek.OnChange += DaysOfTheWeekChanged;
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="WeeklyPattern" /> class.
        /// </summary>
        /// <param name="startDate">The date and time when the recurrence starts.</param>
        /// <param name="interval">The number of weeks between each occurrence.</param>
        /// <param name="daysOfTheWeek">The days of the week when occurrences happen.</param>
        public WeeklyPattern(DateTime startDate, int interval, params DayOfTheWeek[] daysOfTheWeek)
            : base(startDate, interval)
        {
            DaysOfTheWeek.AddRange(daysOfTheWeek);
        }

        /// <summary>
        ///     Change event handler.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void DaysOfTheWeekChanged(ComplexProperty complexProperty)
        {
            Changed();
        }

        /// <summary>
        ///     Write properties to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWritePropertiesToXml(EwsServiceXmlWriter writer)
        {
            base.InternalWritePropertiesToXml(writer);

            DaysOfTheWeek.WriteToXml(writer, XmlElementNames.DaysOfWeek);

            if (_firstDayOfWeek.HasValue)
            {
                //  We only allow the "FirstDayOfWeek" parameter for the Exchange2010_SP1 schema
                //  version.
                //
                EwsUtilities.ValidatePropertyVersion(
                    (ExchangeService)writer.Service,
                    ExchangeVersion.Exchange2010_SP1,
                    nameof(FirstDayOfWeek)
                );

                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FirstDayOfWeek, _firstDayOfWeek.Value);
            }
        }

        /// <summary>
        ///     Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (base.TryReadElementFromXml(reader))
            {
                return true;
            }

            switch (reader.LocalName)
            {
                case XmlElementNames.DaysOfWeek:
                {
                    DaysOfTheWeek.LoadFromXml(reader, reader.LocalName);
                    return true;
                }
                case XmlElementNames.FirstDayOfWeek:
                {
                    FirstDayOfWeek = reader.ReadElementValue<DayOfWeek>(
                        XmlNamespace.Types,
                        XmlElementNames.FirstDayOfWeek
                    );
                    return true;
                }
                default:
                {
                    return false;
                }
            }
        }

        /// <summary>
        ///     Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();

            if (DaysOfTheWeek.Count == 0)
            {
                throw new ServiceValidationException(Strings.DaysOfTheWeekNotSpecified);
            }
        }

        /// <summary>
        ///     Checks if two recurrence objects are identical.
        /// </summary>
        /// <param name="otherRecurrence">The recurrence to compare this one to.</param>
        /// <returns>true if the two recurrences are identical, false otherwise.</returns>
        public override bool IsSame(Recurrence? otherRecurrence)
        {
            var otherWeeklyPattern = (WeeklyPattern?)otherRecurrence;

            return base.IsSame(otherRecurrence) &&
                   DaysOfTheWeek.ToString(",") == otherWeeklyPattern.DaysOfTheWeek.ToString(",") &&
                   _firstDayOfWeek == otherWeeklyPattern._firstDayOfWeek;
        }
    }
}
