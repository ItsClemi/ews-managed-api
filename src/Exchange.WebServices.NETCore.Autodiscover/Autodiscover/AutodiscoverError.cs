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
using System.Xml;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents an error returned by the Autodiscover service.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class AutodiscoverError
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverError" /> class.
    /// </summary>
    private AutodiscoverError()
    {
    }

    /// <summary>
    ///     Parses the XML through the specified reader and creates an Autodiscover error.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>An Autodiscover error.</returns>
    internal static AutodiscoverError Parse(EwsXmlReader reader)
    {
        var error = new AutodiscoverError
        {
            Time = reader.ReadAttributeValue(XmlAttributeNames.Time),
            Id = reader.ReadAttributeValue(XmlAttributeNames.Id),
        };

        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.ErrorCode:
                    {
                        error.ErrorCode = reader.ReadElementValue<int>();
                        break;
                    }
                    case XmlElementNames.Message:
                    {
                        error.Message = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.DebugData:
                    {
                        error.DebugData = reader.ReadElementValue();
                        break;
                    }
                    default:
                    {
                        reader.SkipCurrentElement();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Error));

        return error;
    }

    /// <summary>
    ///     Gets the time when the error was returned.
    /// </summary>
    public string Time { get; private set; }

    /// <summary>
    ///     Gets a hash of the name of the computer that is running Microsoft Exchange Server that has the Client Access server
    ///     role installed.
    /// </summary>
    public string Id { get; private set; }

    /// <summary>
    ///     Gets the error code.
    /// </summary>
    public int ErrorCode { get; private set; }

    /// <summary>
    ///     Gets the error message.
    /// </summary>
    public string Message { get; private set; }

    /// <summary>
    ///     Gets the debug data.
    /// </summary>
    public string DebugData { get; private set; }
}
