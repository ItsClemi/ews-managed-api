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
///     Represents the schema for contact groups.
/// </summary>
[PublicAPI]
[Schema]
public class ContactGroupSchema : ItemSchema
{
    /// <summary>
    ///     Defines the DisplayName property.
    /// </summary>
    public static readonly PropertyDefinition DisplayName = ContactSchema.DisplayName;

    /// <summary>
    ///     Defines the FileAs property.
    /// </summary>
    public static readonly PropertyDefinition FileAs = ContactSchema.FileAs;

    /// <summary>
    ///     Defines the Members property.
    /// </summary>
    public static readonly PropertyDefinition Members = new ComplexPropertyDefinition<GroupMemberCollection>(
        XmlElementNames.Members,
        FieldUris.Members,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate,
        ExchangeVersion.Exchange2010,
        () => new GroupMemberCollection()
    );

    /// <summary>
    ///     This must be declared after the property definitions.
    /// </summary>
    internal new static readonly ContactGroupSchema Instance = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="ContactGroupSchema" /> class.
    /// </summary>
    internal ContactGroupSchema()
    {
    }

    /// <summary>
    ///     Registers properties.
    /// </summary>
    /// <remarks>
    ///     IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in
    ///     types.xsd)
    /// </remarks>
    internal override void RegisterProperties()
    {
        base.RegisterProperties();

        RegisterProperty(DisplayName);
        RegisterProperty(FileAs);
        RegisterProperty(Members);
    }

    /// <summary>
    ///     Field URIs for Members.
    /// </summary>
    private static class FieldUris
    {
        /// <summary>
        ///     FieldUri for members.
        /// </summary>
        public const string Members = "distributionlist:Members";
    }
}
