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
///     Represents a user's Out of Office (OOF) settings.
/// </summary>
[PublicAPI]
public sealed class OofSettings : ComplexProperty, ISelfValidate
{
    /// <summary>
    ///     Gets or sets the user's OOF state.
    /// </summary>
    /// <value>The user's OOF state.</value>
    public OofState State { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating who should receive external OOF messages.
    /// </summary>
    public OofExternalAudience ExternalAudience { get; set; }

    /// <summary>
    ///     Gets or sets the duration of the OOF status when State is set to OofState.Scheduled.
    /// </summary>
    public TimeWindow Duration { get; set; }

    /// <summary>
    ///     Gets or sets the OOF response sent other users in the user's domain or trusted domain.
    /// </summary>
    public OofReply InternalReply { get; set; }

    /// <summary>
    ///     Gets or sets the OOF response sent to addresses outside the user's domain or trusted domain.
    /// </summary>
    public OofReply ExternalReply { get; set; }

    /// <summary>
    ///     Gets a value indicating the authorized external OOF notifications.
    /// </summary>
    public OofExternalAudience AllowExternalOof { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of OofSettings.
    /// </summary>
    public OofSettings()
    {
    }


    #region ISelfValidate Members

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    void ISelfValidate.Validate()
    {
        if (State == OofState.Scheduled)
        {
            if (Duration == null)
            {
                throw new ArgumentException(Strings.DurationMustBeSpecifiedWhenScheduled);
            }

            EwsUtilities.ValidateParam(Duration);
        }
    }

    #endregion


    /// <summary>
    ///     Serializes an OofReply. Emits an empty OofReply in case the one passed in is null.
    /// </summary>
    /// <param name="oofReply">The oof reply.</param>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    private static void SerializeOofReply(OofReply? oofReply, EwsServiceXmlWriter writer, string xmlElementName)
    {
        if (oofReply != null)
        {
            oofReply.WriteToXml(writer, xmlElementName);
        }
        else
        {
            OofReply.WriteEmptyReplyToXml(writer, xmlElementName);
        }
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if appropriate element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.OofState:
            {
                State = reader.ReadValue<OofState>();
                return true;
            }
            case XmlElementNames.ExternalAudience:
            {
                ExternalAudience = reader.ReadValue<OofExternalAudience>();
                return true;
            }
            case XmlElementNames.Duration:
            {
                Duration = new TimeWindow();
                Duration.LoadFromXml(reader);
                return true;
            }
            case XmlElementNames.InternalReply:
            {
                InternalReply = new OofReply();
                InternalReply.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ExternalReply:
            {
                ExternalReply = new OofReply();
                ExternalReply.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        base.WriteElementsToXml(writer);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.OofState, State);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ExternalAudience, ExternalAudience);

        if (Duration != null && State == OofState.Scheduled)
        {
            Duration.WriteToXml(writer, XmlElementNames.Duration);
        }

        SerializeOofReply(InternalReply, writer, XmlElementNames.InternalReply);
        SerializeOofReply(ExternalReply, writer, XmlElementNames.ExternalReply);
    }
}
