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
///     Represents the schema for Conversation.
/// </summary>
[PublicAPI]
[Schema]
public class ConversationSchema : ServiceObjectSchema
{
    /// <summary>
    ///     Field URIs for Item.
    /// </summary>
    private static class FieldUris
    {
        public const string ConversationId = "conversation:ConversationId";
        public const string ConversationTopic = "conversation:ConversationTopic";
        public const string UniqueRecipients = "conversation:UniqueRecipients";
        public const string GlobalUniqueRecipients = "conversation:GlobalUniqueRecipients";
        public const string UniqueUnreadSenders = "conversation:UniqueUnreadSenders";
        public const string GlobalUniqueUnreadSenders = "conversation:GlobalUniqueUnreadSenders";
        public const string UniqueSenders = "conversation:UniqueSenders";
        public const string GlobalUniqueSenders = "conversation:GlobalUniqueSenders";
        public const string LastDeliveryTime = "conversation:LastDeliveryTime";
        public const string GlobalLastDeliveryTime = "conversation:GlobalLastDeliveryTime";
        public const string Categories = "conversation:Categories";
        public const string GlobalCategories = "conversation:GlobalCategories";
        public const string FlagStatus = "conversation:FlagStatus";
        public const string GlobalFlagStatus = "conversation:GlobalFlagStatus";
        public const string HasAttachments = "conversation:HasAttachments";
        public const string GlobalHasAttachments = "conversation:GlobalHasAttachments";
        public const string MessageCount = "conversation:MessageCount";
        public const string GlobalMessageCount = "conversation:GlobalMessageCount";
        public const string UnreadCount = "conversation:UnreadCount";
        public const string GlobalUnreadCount = "conversation:GlobalUnreadCount";
        public const string Size = "conversation:Size";
        public const string GlobalSize = "conversation:GlobalSize";
        public const string ItemClasses = "conversation:ItemClasses";
        public const string GlobalItemClasses = "conversation:GlobalItemClasses";
        public const string Importance = "conversation:Importance";
        public const string GlobalImportance = "conversation:GlobalImportance";
        public const string ItemIds = "conversation:ItemIds";
        public const string GlobalItemIds = "conversation:GlobalItemIds";
        public const string LastModifiedTime = "conversation:LastModifiedTime";
        public const string InstanceKey = "conversation:InstanceKey";
        public const string Preview = "conversation:Preview";
        public const string IconIndex = "conversation:IconIndex";
        public const string GlobalIconIndex = "conversation:GlobalIconIndex";
        public const string DraftItemIds = "conversation:DraftItemIds";
        public const string HasIrm = "conversation:HasIrm";
        public const string GlobalHasIrm = "conversation:GlobalHasIrm";
    }

    /// <summary>
    ///     Defines the Id property.
    /// </summary>
    public static readonly PropertyDefinition Id = new ComplexPropertyDefinition<ConversationId>(
        XmlElementNames.ConversationId,
        FieldUris.ConversationId,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new ConversationId()
    );

    /// <summary>
    ///     Defines the Topic property.
    /// </summary>
    public static readonly PropertyDefinition Topic = new StringPropertyDefinition(
        XmlElementNames.ConversationTopic,
        FieldUris.ConversationTopic,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the UniqueRecipients property.
    /// </summary>
    public static readonly PropertyDefinition UniqueRecipients = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.UniqueRecipients,
        FieldUris.UniqueRecipients,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the GlobalUniqueRecipients property.
    /// </summary>
    public static readonly PropertyDefinition GlobalUniqueRecipients = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.GlobalUniqueRecipients,
        FieldUris.GlobalUniqueRecipients,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the UniqueUnreadSenders property.
    /// </summary>
    public static readonly PropertyDefinition UniqueUnreadSenders = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.UniqueUnreadSenders,
        FieldUris.UniqueUnreadSenders,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the GlobalUniqueUnreadSenders property.
    /// </summary>
    public static readonly PropertyDefinition GlobalUniqueUnreadSenders = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.GlobalUniqueUnreadSenders,
        FieldUris.GlobalUniqueUnreadSenders,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the UniqueSenders property.
    /// </summary>
    public static readonly PropertyDefinition UniqueSenders = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.UniqueSenders,
        FieldUris.UniqueSenders,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the GlobalUniqueSenders property.
    /// </summary>
    public static readonly PropertyDefinition GlobalUniqueSenders = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.GlobalUniqueSenders,
        FieldUris.GlobalUniqueSenders,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the LastDeliveryTime property.
    /// </summary>
    public static readonly PropertyDefinition LastDeliveryTime = new DateTimePropertyDefinition(
        XmlElementNames.LastDeliveryTime,
        FieldUris.LastDeliveryTime,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalLastDeliveryTime property.
    /// </summary>
    public static readonly PropertyDefinition GlobalLastDeliveryTime = new DateTimePropertyDefinition(
        XmlElementNames.GlobalLastDeliveryTime,
        FieldUris.GlobalLastDeliveryTime,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the Categories property.
    /// </summary>
    public static readonly PropertyDefinition Categories = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.Categories,
        FieldUris.Categories,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the GlobalCategories property.
    /// </summary>
    public static readonly PropertyDefinition GlobalCategories = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.GlobalCategories,
        FieldUris.GlobalCategories,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the FlagStatus property.
    /// </summary>
    public static readonly PropertyDefinition FlagStatus = new GenericPropertyDefinition<ConversationFlagStatus>(
        XmlElementNames.FlagStatus,
        FieldUris.FlagStatus,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalFlagStatus property.
    /// </summary>
    public static readonly PropertyDefinition GlobalFlagStatus = new GenericPropertyDefinition<ConversationFlagStatus>(
        XmlElementNames.GlobalFlagStatus,
        FieldUris.GlobalFlagStatus,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the HasAttachments property.
    /// </summary>
    public static readonly PropertyDefinition HasAttachments = new BoolPropertyDefinition(
        XmlElementNames.HasAttachments,
        FieldUris.HasAttachments,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalHasAttachments property.
    /// </summary>
    public static readonly PropertyDefinition GlobalHasAttachments = new BoolPropertyDefinition(
        XmlElementNames.GlobalHasAttachments,
        FieldUris.GlobalHasAttachments,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the MessageCount property.
    /// </summary>
    public static readonly PropertyDefinition MessageCount = new IntPropertyDefinition(
        XmlElementNames.MessageCount,
        FieldUris.MessageCount,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalMessageCount property.
    /// </summary>
    public static readonly PropertyDefinition GlobalMessageCount = new IntPropertyDefinition(
        XmlElementNames.GlobalMessageCount,
        FieldUris.GlobalMessageCount,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the UnreadCount property.
    /// </summary>
    public static readonly PropertyDefinition UnreadCount = new IntPropertyDefinition(
        XmlElementNames.UnreadCount,
        FieldUris.UnreadCount,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalUnreadCount property.
    /// </summary>
    public static readonly PropertyDefinition GlobalUnreadCount = new IntPropertyDefinition(
        XmlElementNames.GlobalUnreadCount,
        FieldUris.GlobalUnreadCount,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the Size property.
    /// </summary>
    public static readonly PropertyDefinition Size = new IntPropertyDefinition(
        XmlElementNames.Size,
        FieldUris.Size,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalSize property.
    /// </summary>
    public static readonly PropertyDefinition GlobalSize = new IntPropertyDefinition(
        XmlElementNames.GlobalSize,
        FieldUris.GlobalSize,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the ItemClasses property.
    /// </summary>
    public static readonly PropertyDefinition ItemClasses = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.ItemClasses,
        FieldUris.ItemClasses,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList(XmlElementNames.ItemClass)
    );

    /// <summary>
    ///     Defines the GlobalItemClasses property.
    /// </summary>
    public static readonly PropertyDefinition GlobalItemClasses = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.GlobalItemClasses,
        FieldUris.GlobalItemClasses,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new StringList(XmlElementNames.ItemClass)
    );

    /// <summary>
    ///     Defines the Importance property.
    /// </summary>
    public static readonly PropertyDefinition Importance = new GenericPropertyDefinition<Importance>(
        XmlElementNames.Importance,
        FieldUris.Importance,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the GlobalImportance property.
    /// </summary>
    public static readonly PropertyDefinition GlobalImportance = new GenericPropertyDefinition<Importance>(
        XmlElementNames.GlobalImportance,
        FieldUris.GlobalImportance,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1
    );

    /// <summary>
    ///     Defines the ItemIds property.
    /// </summary>
    public static readonly PropertyDefinition ItemIds = new ComplexPropertyDefinition<ItemIdCollection>(
        XmlElementNames.ItemIds,
        FieldUris.ItemIds,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new ItemIdCollection()
    );

    /// <summary>
    ///     Defines the GlobalItemIds property.
    /// </summary>
    public static readonly PropertyDefinition GlobalItemIds = new ComplexPropertyDefinition<ItemIdCollection>(
        XmlElementNames.GlobalItemIds,
        FieldUris.GlobalItemIds,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP1,
        () => new ItemIdCollection()
    );

    /// <summary>
    ///     Defines the LastModifiedTime property.
    /// </summary>
    public static readonly PropertyDefinition LastModifiedTime = new DateTimePropertyDefinition(
        XmlElementNames.LastModifiedTime,
        FieldUris.LastModifiedTime,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the InstanceKey property.
    /// </summary>
    public static readonly PropertyDefinition InstanceKey = new ByteArrayPropertyDefinition(
        XmlElementNames.InstanceKey,
        FieldUris.InstanceKey,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the Preview property.
    /// </summary>
    public static readonly PropertyDefinition Preview = new StringPropertyDefinition(
        XmlElementNames.Preview,
        FieldUris.Preview,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the IconIndex property.
    /// </summary>
    public static readonly PropertyDefinition IconIndex = new GenericPropertyDefinition<IconIndex>(
        XmlElementNames.IconIndex,
        FieldUris.IconIndex,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the GlobalIconIndex property.
    /// </summary>
    public static readonly PropertyDefinition GlobalIconIndex = new GenericPropertyDefinition<IconIndex>(
        XmlElementNames.GlobalIconIndex,
        FieldUris.GlobalIconIndex,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the DraftItemIds property.
    /// </summary>
    public static readonly PropertyDefinition DraftItemIds = new ComplexPropertyDefinition<ItemIdCollection>(
        XmlElementNames.DraftItemIds,
        FieldUris.DraftItemIds,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        () => new ItemIdCollection()
    );

    /// <summary>
    ///     Defines the HasIrm property.
    /// </summary>
    public static readonly PropertyDefinition HasIrm = new BoolPropertyDefinition(
        XmlElementNames.HasIrm,
        FieldUris.HasIrm,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    /// <summary>
    ///     Defines the GlobalHasIrm property.
    /// </summary>
    public static readonly PropertyDefinition GlobalHasIrm = new BoolPropertyDefinition(
        XmlElementNames.GlobalHasIrm,
        FieldUris.GlobalHasIrm,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    // This must be declared after the property definitions
    internal static readonly ConversationSchema Instance = new();

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

        RegisterProperty(Id);
        RegisterProperty(Topic);
        RegisterProperty(UniqueRecipients);
        RegisterProperty(GlobalUniqueRecipients);
        RegisterProperty(UniqueUnreadSenders);
        RegisterProperty(GlobalUniqueUnreadSenders);
        RegisterProperty(UniqueSenders);
        RegisterProperty(GlobalUniqueSenders);
        RegisterProperty(LastDeliveryTime);
        RegisterProperty(GlobalLastDeliveryTime);
        RegisterProperty(Categories);
        RegisterProperty(GlobalCategories);
        RegisterProperty(FlagStatus);
        RegisterProperty(GlobalFlagStatus);
        RegisterProperty(HasAttachments);
        RegisterProperty(GlobalHasAttachments);
        RegisterProperty(MessageCount);
        RegisterProperty(GlobalMessageCount);
        RegisterProperty(UnreadCount);
        RegisterProperty(GlobalUnreadCount);
        RegisterProperty(Size);
        RegisterProperty(GlobalSize);
        RegisterProperty(ItemClasses);
        RegisterProperty(GlobalItemClasses);
        RegisterProperty(Importance);
        RegisterProperty(GlobalImportance);
        RegisterProperty(ItemIds);
        RegisterProperty(GlobalItemIds);
        RegisterProperty(LastModifiedTime);
        RegisterProperty(InstanceKey);
        RegisterProperty(Preview);
        RegisterProperty(IconIndex);
        RegisterProperty(GlobalIconIndex);
        RegisterProperty(DraftItemIds);
        RegisterProperty(HasIrm);
        RegisterProperty(GlobalHasIrm);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConversationSchema" /> class.
    /// </summary>
    internal ConversationSchema()
    {
    }
}
