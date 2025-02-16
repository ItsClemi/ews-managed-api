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

/// <content>
///     Contains nested type SearchFilter.RelationalFilter.
/// </content>
public abstract partial class SearchFilter
{
    /// <summary>
    ///     Represents the base class for relational filters (for example, IsEqualTo, IsGreaterThan or IsLessThanOrEqualTo).
    /// </summary>
    [PublicAPI]
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class RelationalFilter : PropertyBasedFilter
    {
        private PropertyDefinitionBase? _otherPropertyDefinition;
        private object? _value;

        /// <summary>
        ///     Gets or sets the definition of the property to compare with. Property definitions are available as static members
        ///     from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName,
        ///     etc.)
        ///     The OtherPropertyDefinition and Value properties are mutually exclusive; setting one resets the other to null.
        /// </summary>
        public PropertyDefinitionBase? OtherPropertyDefinition
        {
            get => _otherPropertyDefinition;

            set
            {
                SetFieldValue(ref _otherPropertyDefinition, value);
                _value = null;
            }
        }

        /// <summary>
        ///     Gets or sets the value to compare with. The Value and OtherPropertyDefinition properties
        ///     are mutually exclusive; setting one resets the other to null.
        /// </summary>
        public object? Value
        {
            get => _value;

            set
            {
                SetFieldValue(ref _value, value);
                _otherPropertyDefinition = null;
            }
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SearchFilter.RelationalFilter" /> class.
        /// </summary>
        internal RelationalFilter()
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SearchFilter.RelationalFilter" /> class.
        /// </summary>
        /// <param name="propertyDefinition">
        ///     The definition of the property that is being compared. Property definitions are
        ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
        ///     ContactSchema.GivenName, etc.)
        /// </param>
        /// <param name="otherPropertyDefinition">
        ///     The definition of the property to compare with. Property definitions are
        ///     available as static members from schema classes (for example, EmailMessageSchema, AppointmentSchema, etc.)
        /// </param>
        internal RelationalFilter(
            PropertyDefinitionBase propertyDefinition,
            PropertyDefinitionBase otherPropertyDefinition
        )
            : base(propertyDefinition)
        {
            _otherPropertyDefinition = otherPropertyDefinition;
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SearchFilter.RelationalFilter" /> class.
        /// </summary>
        /// <param name="propertyDefinition">
        ///     The definition of the property that is being compared. Property definitions are
        ///     available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start,
        ///     ContactSchema.GivenName, etc.)
        /// </param>
        /// <param name="value">The value to compare with.</param>
        internal RelationalFilter(PropertyDefinitionBase propertyDefinition, object value)
            : base(propertyDefinition)
        {
            _value = value;
        }

        /// <summary>
        ///     Validate instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();

            if (_otherPropertyDefinition == null && _value == null)
            {
                throw new ServiceValidationException(Strings.EqualityComparisonFilterIsInvalid);
            }

            if (_value != null)
            {
                // All common value types (String, Int32, DateTime, ...) implement IConvertible.
                // Value types that don't implement IConvertible must implement ISearchStringProvider 
                // in order to be used in a search filter.
                if (!(_value is IConvertible || _value is ISearchStringProvider))
                {
                    throw new ServiceValidationException(
                        string.Format(Strings.SearchFilterComparisonValueTypeIsNotSupported, _value.GetType().Name)
                    );
                }
            }
        }

        /// <summary>
        ///     Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            var result = base.TryReadElementFromXml(reader);

            if (!result)
            {
                if (reader.LocalName == XmlElementNames.FieldURIOrConstant)
                {
                    reader.Read();
                    reader.EnsureCurrentNodeIsStartElement();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Constant))
                    {
                        _value = reader.ReadAttributeValue(XmlAttributeNames.Value);

                        result = true;
                    }
                    else
                    {
                        result = PropertyDefinitionBase.TryLoadFromXml(reader, ref _otherPropertyDefinition);
                    }
                }
            }

            return result;
        }

        /// <summary>
        ///     Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FieldURIOrConstant);

            if (Value != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Constant);
                writer.WriteAttributeValue(XmlAttributeNames.Value, true /* alwaysWriteEmptyString */, Value);
                writer.WriteEndElement(); // Constant
            }
            else
            {
                OtherPropertyDefinition.WriteToXml(writer);
            }

            writer.WriteEndElement(); // FieldURIOrConstant
        }
    }
}
