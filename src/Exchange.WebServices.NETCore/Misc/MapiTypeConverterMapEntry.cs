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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an entry in the MapiTypeConverter map.
/// </summary>
internal class MapiTypeConverterMapEntry
{
    /// <summary>
    ///     Map CLR types used for MAPI properties to matching default values.
    /// </summary>
    private static readonly IReadOnlyDictionary<Type, object?> DefaultValueMap = new Dictionary<Type, object?>
    {
        // @formatter:off
        { typeof(bool), false },
        { typeof(byte[]), null },
        { typeof(short), (short)0 },
        { typeof(int), 0 },
        { typeof(long), (long)0 },
        { typeof(float), 0.0f },
        { typeof(double), 0.0 },
        { typeof(DateTime), DateTime.MinValue },
        { typeof(Guid), Guid.Empty },
        { typeof(string), null },
        // @formatter:on
    };

    /// <summary>
    ///     Initializes a new instance of the <see cref="MapiTypeConverterMapEntry" /> class.
    /// </summary>
    /// <param name="type">The type.</param>
    /// <remarks>
    ///     By default, converting a type to string is done by calling value.ToString. Instances
    ///     can override this behavior.
    ///     By default, converting a string to the appropriate value type is done by calling Convert.ChangeType
    ///     Instances may override this behavior.
    /// </remarks>
    internal MapiTypeConverterMapEntry(Type type)
    {
        EwsUtilities.Assert(
            DefaultValueMap.ContainsKey(type),
            "MapiTypeConverterMapEntry ctor",
            $"No default value entry for type {type.Name}"
        );

        Type = type;
        ConvertToString = o => (string)Convert.ChangeType(o, typeof(string), CultureInfo.InvariantCulture);
        Parse = s => Convert.ChangeType(s, type, CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Change value to a value of compatible type.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>New value.</returns>
    /// <remarks>
    ///     The type of a simple value should match exactly or be convertible to the appropriate type. An
    ///     array value has to be a single dimension (rank), contain at least one value and contain
    ///     elements that exactly match the expected type. (We could relax this last requirement so that,
    ///     for example, you could pass an array of Int32 that could be converted to an array of Double
    ///     but that seems like overkill).
    /// </remarks>
    internal object ChangeType(object value)
    {
        if (IsArray)
        {
            ValidateValueAsArray(value);
            return value;
        }

        if (value.GetType() == Type)
        {
            return value;
        }

        try
        {
            return Convert.ChangeType(value, Type, CultureInfo.InvariantCulture);
        }
        catch (InvalidCastException ex)
        {
            throw new ArgumentException(
                string.Format(Strings.ValueOfTypeCannotBeConverted, value, value.GetType(), Type),
                nameof(value),
                ex
            );
        }
    }

    /// <summary>
    ///     Converts a string to value consistent with type.
    /// </summary>
    /// <param name="stringValue">String to convert to a value.</param>
    /// <returns>Value.</returns>
    internal object? ConvertToValue(string stringValue)
    {
        try
        {
            return Parse(stringValue);
        }
        catch (FormatException ex)
        {
            throw new ServiceXmlDeserializationException(
                string.Format(Strings.ValueCannotBeConverted, stringValue, Type),
                ex
            );
        }
        catch (InvalidCastException ex)
        {
            throw new ServiceXmlDeserializationException(
                string.Format(Strings.ValueCannotBeConverted, stringValue, Type),
                ex
            );
        }
        catch (OverflowException ex)
        {
            throw new ServiceXmlDeserializationException(
                string.Format(Strings.ValueCannotBeConverted, stringValue, Type),
                ex
            );
        }
    }

    /// <summary>
    ///     Converts a string to value consistent with type (or uses the default value if the string is null or empty).
    /// </summary>
    /// <param name="stringValue">String to convert to a value.</param>
    /// <returns>Value.</returns>
    /// <remarks>For array types, this method is called for each array element.</remarks>
    internal object? ConvertToValueOrDefault(string stringValue)
    {
        return string.IsNullOrEmpty(stringValue) ? DefaultValue : ConvertToValue(stringValue);
    }

    /// <summary>
    ///     Validates array value.
    /// </summary>
    /// <param name="value">The value.</param>
    private void ValidateValueAsArray(object value)
    {
        if (value is not Array array)
        {
            throw new ArgumentException(
                string.Format(Strings.IncompatibleTypeForArray, value.GetType(), Type),
                nameof(value)
            );
        }

        if (array.Rank != 1)
        {
            throw new ArgumentException(Strings.ArrayMustHaveSingleDimension, nameof(value));
        }

        if (array.Length == 0)
        {
            throw new ArgumentException(Strings.ArrayMustHaveAtLeastOneElement, nameof(value));
        }

        if (array.GetType().GetElementType() != Type)
        {
            throw new ArgumentException(
                string.Format(Strings.IncompatibleTypeForArray, value.GetType(), Type),
                nameof(value)
            );
        }
    }


    #region Properties

    /// <summary>
    ///     Gets or sets the string parser.
    /// </summary>
    /// <remarks>For array types, this method is called for each array element.</remarks>
    internal Func<string?, object?> Parse { get; set; }

    /// <summary>
    ///     Gets or sets the string to object converter.
    /// </summary>
    /// <remarks>For array types, this method is called for each array element.</remarks>
    internal Func<object, string> ConvertToString { get; set; }

    /// <summary>
    ///     Gets or sets the type.
    /// </summary>
    /// <remarks>For array types, this is the type of an element.</remarks>
    internal Type Type { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether this instance is array.
    /// </summary>
    /// <value><c>true</c> if this instance is array; otherwise, <c>false</c>.</value>
    internal bool IsArray { get; set; }

    /// <summary>
    ///     Gets the default value for the type.
    /// </summary>
    internal object? DefaultValue => DefaultValueMap[Type];

    #endregion
}
