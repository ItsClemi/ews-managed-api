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

using System.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the MultiValueInsightContent.
/// </summary>
[PublicAPI]
public sealed class MultiValueInsightContent : ComplexProperty
{
    /// <summary>
    ///     Gets the ItemList
    /// </summary>
    public InsightValueCollection ItemList { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MultiValueInsightContent" /> class.
    /// </summary>
    public MultiValueInsightContent()
    {
        ItemList = new InsightValueCollection();
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XmlElementNames.Item)
            {
                switch (reader.ReadAttributeValue("xsi:type"))
                {
                    case XmlElementNames.StringInsightValue:
                    {
                        var item = new StringInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.ProfileInsightValue:
                    {
                        var item = new ProfileInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.JobInsightValue:
                    {
                        var item = new JobInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.UserProfilePicture:
                    {
                        var item = new UserProfilePicture();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.EducationInsightValue:
                    {
                        var item = new EducationInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.SkillInsightValue:
                    {
                        var item = new SkillInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.ComputedInsightValue:
                    {
                        var item = new ComputedInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.MeetingInsightValue:
                    {
                        var item = new MeetingInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.EmailInsightValue:
                    {
                        var item = new EmailInsightValue();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    case XmlElementNames.DelveDocument:
                    {
                        var item = new DelveDocument();
                        item.LoadFromXml(reader, reader.LocalName);
                        ItemList.InternalAdd(item);
                        break;
                    }
                    default:
                    {
                        return false;
                    }
                }
            }
            else
            {
                return false;
            }
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ItemList));

        return true;
    }
}
