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
///     Represents voting option information.
/// </summary>
[PublicAPI]
public sealed class VotingOptionData : ComplexProperty
{
    /// <summary>
    ///     Gets the display name for the voting option.
    /// </summary>
    public string DisplayName { get; private set; }

    /// <summary>
    ///     Gets the send prompt.
    /// </summary>
    public SendPrompt SendPrompt { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="VotingOptionData" /> class.
    /// </summary>
    internal VotingOptionData()
    {
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.VotingOptionDisplayName:
            {
                DisplayName = reader.ReadElementValue<string>();
                return true;
            }
            case XmlElementNames.SendPrompt:
            {
                SendPrompt = reader.ReadElementValue<SendPrompt>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
