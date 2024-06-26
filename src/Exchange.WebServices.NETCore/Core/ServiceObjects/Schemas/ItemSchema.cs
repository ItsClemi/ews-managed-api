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
///     Represents the schema for generic items.
/// </summary>
[PublicAPI]
[Schema]
public class ItemSchema : ServiceObjectSchema
{
    /// <summary>
    ///     Field URIs for Item.
    /// </summary>
    private static class FieldUris
    {
        public const string ItemId = "item:ItemId";
        public const string ParentFolderId = "item:ParentFolderId";
        public const string ItemClass = "item:ItemClass";
        public const string MimeContent = "item:MimeContent";
        public const string MimeContentUTF8 = "item:MimeContentUTF8";
        public const string Attachments = "item:Attachments";
        public const string Subject = "item:Subject";
        public const string DateTimeReceived = "item:DateTimeReceived";
        public const string Size = "item:Size";
        public const string Categories = "item:Categories";
        public const string HasAttachments = "item:HasAttachments";
        public const string Importance = "item:Importance";
        public const string InReplyTo = "item:InReplyTo";
        public const string InternetMessageHeaders = "item:InternetMessageHeaders";
        public const string IsAssociated = "item:IsAssociated";
        public const string IsDraft = "item:IsDraft";
        public const string IsFromMe = "item:IsFromMe";
        public const string IsResend = "item:IsResend";
        public const string IsSubmitted = "item:IsSubmitted";
        public const string IsUnmodified = "item:IsUnmodified";
        public const string DateTimeSent = "item:DateTimeSent";
        public const string DateTimeCreated = "item:DateTimeCreated";
        public const string Body = "item:Body";
        public const string ResponseObjects = "item:ResponseObjects";
        public const string Sensitivity = "item:Sensitivity";
        public const string ReminderDueBy = "item:ReminderDueBy";
        public const string ReminderIsSet = "item:ReminderIsSet";
        public const string ReminderMinutesBeforeStart = "item:ReminderMinutesBeforeStart";
        public const string DisplayTo = "item:DisplayTo";
        public const string DisplayCc = "item:DisplayCc";
        public const string Culture = "item:Culture";
        public const string EffectiveRights = "item:EffectiveRights";
        public const string LastModifiedName = "item:LastModifiedName";
        public const string LastModifiedTime = "item:LastModifiedTime";
        public const string WebClientReadFormQueryString = "item:WebClientReadFormQueryString";
        public const string WebClientEditFormQueryString = "item:WebClientEditFormQueryString";
        public const string ConversationId = "item:ConversationId";
        public const string UniqueBody = "item:UniqueBody";
        public const string StoreEntryId = "item:StoreEntryId";
        public const string InstanceKey = "item:InstanceKey";
        public const string NormalizedBody = "item:NormalizedBody";
        public const string EntityExtractionResult = "item:EntityExtractionResult";
        public const string Flag = "item:Flag";
        public const string PolicyTag = "item:PolicyTag";
        public const string ArchiveTag = "item:ArchiveTag";
        public const string RetentionDate = "item:RetentionDate";
        public const string Preview = "item:Preview";
        public const string TextBody = "item:TextBody";
        public const string IconIndex = "item:IconIndex";
        public const string Hashtags = "item:Hashtags";
        public const string Mentions = "item:Mentions";
        public const string MentionedMe = "item:MentionedMe";
    }

    /// <summary>
    ///     Defines the Id property.
    /// </summary>
    public static readonly PropertyDefinition Id = new ComplexPropertyDefinition<ItemId>(
        XmlElementNames.ItemId,
        FieldUris.ItemId,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        () => new ItemId()
    );

    /// <summary>
    ///     Defines the Body property.
    /// </summary>
    public static readonly PropertyDefinition Body = new ComplexPropertyDefinition<MessageBody>(
        XmlElementNames.Body,
        FieldUris.Body,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
        ExchangeVersion.Exchange2007_SP1,
        () => new MessageBody()
    );

    /// <summary>
    ///     Defines the ItemClass property.
    /// </summary>
    public static readonly PropertyDefinition ItemClass = new StringPropertyDefinition(
        XmlElementNames.ItemClass,
        FieldUris.ItemClass,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Subject property.
    /// </summary>
    public static readonly PropertyDefinition Subject = new StringPropertyDefinition(
        XmlElementNames.Subject,
        FieldUris.Subject,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the MimeContent property.
    /// </summary>
    public static readonly PropertyDefinition MimeContent = new ComplexPropertyDefinition<MimeContent>(
        XmlElementNames.MimeContent,
        FieldUris.MimeContent,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.MustBeExplicitlyLoaded,
        ExchangeVersion.Exchange2007_SP1,
        () => new MimeContent()
    );

    /// <summary>
    ///     Defines the MimeContentUTF8 property.
    /// </summary>
    public static readonly PropertyDefinition MimeContentUTF8 = new ComplexPropertyDefinition<MimeContentUTF8>(
        XmlElementNames.MimeContentUTF8,
        FieldUris.MimeContentUTF8,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.MustBeExplicitlyLoaded,
        ExchangeVersion.Exchange2013_SP1,
        () => new MimeContentUTF8()
    );

    /// <summary>
    ///     Defines the ParentFolderId property.
    /// </summary>
    public static readonly PropertyDefinition ParentFolderId = new ComplexPropertyDefinition<FolderId>(
        XmlElementNames.ParentFolderId,
        FieldUris.ParentFolderId,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        () => new FolderId()
    );

    /// <summary>
    ///     Defines the Sensitivity property.
    /// </summary>
    public static readonly PropertyDefinition Sensitivity = new GenericPropertyDefinition<Sensitivity>(
        XmlElementNames.Sensitivity,
        FieldUris.Sensitivity,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Attachments property.
    /// </summary>
    public static readonly PropertyDefinition Attachments = new AttachmentsPropertyDefinition();

    /// <summary>
    ///     Defines the DateTimeReceived property.
    /// </summary>
    public static readonly PropertyDefinition DateTimeReceived = new DateTimePropertyDefinition(
        XmlElementNames.DateTimeReceived,
        FieldUris.DateTimeReceived,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Size property.
    /// </summary>
    public static readonly PropertyDefinition Size = new IntPropertyDefinition(
        XmlElementNames.Size,
        FieldUris.Size,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Categories property.
    /// </summary>
    public static readonly PropertyDefinition Categories = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.Categories,
        FieldUris.Categories,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the Importance property.
    /// </summary>
    public static readonly PropertyDefinition Importance = new GenericPropertyDefinition<Importance>(
        XmlElementNames.Importance,
        FieldUris.Importance,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the InReplyTo property.
    /// </summary>
    public static readonly PropertyDefinition InReplyTo = new StringPropertyDefinition(
        XmlElementNames.InReplyTo,
        FieldUris.InReplyTo,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsSubmitted property.
    /// </summary>
    public static readonly PropertyDefinition IsSubmitted = new BoolPropertyDefinition(
        XmlElementNames.IsSubmitted,
        FieldUris.IsSubmitted,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsAssociated property.
    /// </summary>
    public static readonly PropertyDefinition IsAssociated = new BoolPropertyDefinition(
        XmlElementNames.IsAssociated,
        FieldUris.IsAssociated,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010
    );

    /// <summary>
    ///     Defines the IsDraft property.
    /// </summary>
    public static readonly PropertyDefinition IsDraft = new BoolPropertyDefinition(
        XmlElementNames.IsDraft,
        FieldUris.IsDraft,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsFromMe property.
    /// </summary>
    public static readonly PropertyDefinition IsFromMe = new BoolPropertyDefinition(
        XmlElementNames.IsFromMe,
        FieldUris.IsFromMe,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsResend property.
    /// </summary>
    public static readonly PropertyDefinition IsResend = new BoolPropertyDefinition(
        XmlElementNames.IsResend,
        FieldUris.IsResend,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsUnmodified property.
    /// </summary>
    public static readonly PropertyDefinition IsUnmodified = new BoolPropertyDefinition(
        XmlElementNames.IsUnmodified,
        FieldUris.IsUnmodified,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the InternetMessageHeaders property.
    /// </summary>
    public static readonly PropertyDefinition InternetMessageHeaders =
        new ComplexPropertyDefinition<InternetMessageHeaderCollection>(
            XmlElementNames.InternetMessageHeaders,
            FieldUris.InternetMessageHeaders,
            ExchangeVersion.Exchange2007_SP1,
            () => new InternetMessageHeaderCollection()
        );

    /// <summary>
    ///     Defines the DateTimeSent property.
    /// </summary>
    public static readonly PropertyDefinition DateTimeSent = new DateTimePropertyDefinition(
        XmlElementNames.DateTimeSent,
        FieldUris.DateTimeSent,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the DateTimeCreated property.
    /// </summary>
    public static readonly PropertyDefinition DateTimeCreated = new DateTimePropertyDefinition(
        XmlElementNames.DateTimeCreated,
        FieldUris.DateTimeCreated,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the AllowedResponseActions property.
    /// </summary>
    public static readonly PropertyDefinition AllowedResponseActions = new ResponseObjectsPropertyDefinition(
        XmlElementNames.ResponseObjects,
        FieldUris.ResponseObjects,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the ReminderDueBy property.
    /// </summary>
    public static readonly PropertyDefinition ReminderDueBy = new ScopedDateTimePropertyDefinition(
        XmlElementNames.ReminderDueBy,
        FieldUris.ReminderDueBy,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        _ => AppointmentSchema.StartTimeZone
    );

    /// <summary>
    ///     Defines the IsReminderSet property.
    /// </summary>
    public static readonly PropertyDefinition IsReminderSet = new BoolPropertyDefinition(
        XmlElementNames.ReminderIsSet, // Note: server-side the name is ReminderIsSet
        FieldUris.ReminderIsSet,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the ReminderMinutesBeforeStart property.
    /// </summary>
    public static readonly PropertyDefinition ReminderMinutesBeforeStart = new IntPropertyDefinition(
        XmlElementNames.ReminderMinutesBeforeStart,
        FieldUris.ReminderMinutesBeforeStart,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the DisplayCc property.
    /// </summary>
    public static readonly PropertyDefinition DisplayCc = new StringPropertyDefinition(
        XmlElementNames.DisplayCc,
        FieldUris.DisplayCc,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the DisplayTo property.
    /// </summary>
    public static readonly PropertyDefinition DisplayTo = new StringPropertyDefinition(
        XmlElementNames.DisplayTo,
        FieldUris.DisplayTo,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the HasAttachments property.
    /// </summary>
    public static readonly PropertyDefinition HasAttachments = new BoolPropertyDefinition(
        XmlElementNames.HasAttachments,
        FieldUris.HasAttachments,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Culture property.
    /// </summary>
    public static readonly PropertyDefinition Culture = new StringPropertyDefinition(
        XmlElementNames.Culture,
        FieldUris.Culture,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the EffectiveRights property.
    /// </summary>
    public static readonly PropertyDefinition EffectiveRights = new EffectiveRightsPropertyDefinition(
        XmlElementNames.EffectiveRights,
        FieldUris.EffectiveRights,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the LastModifiedName property.
    /// </summary>
    public static readonly PropertyDefinition LastModifiedName = new StringPropertyDefinition(
        XmlElementNames.LastModifiedName,
        FieldUris.LastModifiedName,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the LastModifiedTime property.
    /// </summary>
    public static readonly PropertyDefinition LastModifiedTime = new DateTimePropertyDefinition(
        XmlElementNames.LastModifiedTime,
        FieldUris.LastModifiedTime,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the WebClientReadFormQueryString property.
    /// </summary>
    public static readonly PropertyDefinition WebClientReadFormQueryString = new StringPropertyDefinition(
        XmlElementNames.WebClientReadFormQueryString,
        FieldUris.WebClientReadFormQueryString,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010
    );

    /// <summary>
    ///     Defines the WebClientEditFormQueryString property.
    /// </summary>
    public static readonly PropertyDefinition WebClientEditFormQueryString = new StringPropertyDefinition(
        XmlElementNames.WebClientEditFormQueryString,
        FieldUris.WebClientEditFormQueryString,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010
    );

    /// <summary>
    ///     Defines the ConversationId property.
    /// </summary>
    public static readonly PropertyDefinition ConversationId = new ComplexPropertyDefinition<ConversationId>(
        XmlElementNames.ConversationId,
        FieldUris.ConversationId,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010,
        () => new ConversationId()
    );

    /// <summary>
    ///     Defines the UniqueBody property.
    /// </summary>
    public static readonly PropertyDefinition UniqueBody = new ComplexPropertyDefinition<UniqueBody>(
        XmlElementNames.UniqueBody,
        FieldUris.UniqueBody,
        PropertyDefinitionFlags.MustBeExplicitlyLoaded,
        ExchangeVersion.Exchange2010,
        () => new UniqueBody()
    );

    /// <summary>
    ///     Defines the StoreEntryId property.
    /// </summary>
    public static readonly PropertyDefinition StoreEntryId = new ByteArrayPropertyDefinition(
        XmlElementNames.StoreEntryId,
        FieldUris.StoreEntryId,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2010_SP2
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
    ///     Defines the NormalizedBody property.
    /// </summary>
    public static readonly PropertyDefinition NormalizedBody = new ComplexPropertyDefinition<NormalizedBody>(
        XmlElementNames.NormalizedBody,
        FieldUris.NormalizedBody,
        PropertyDefinitionFlags.MustBeExplicitlyLoaded,
        ExchangeVersion.Exchange2013,
        () => new NormalizedBody()
    );

    /// <summary>
    ///     Defines the EntityExtractionResult property.
    /// </summary>
    public static readonly PropertyDefinition EntityExtractionResult =
        new ComplexPropertyDefinition<EntityExtractionResult>(
            XmlElementNames.NlgEntityExtractionResult,
            FieldUris.EntityExtractionResult,
            PropertyDefinitionFlags.MustBeExplicitlyLoaded,
            ExchangeVersion.Exchange2013,
            () => new EntityExtractionResult()
        );

    /// <summary>
    ///     Defines the InternetMessageHeaders property.
    /// </summary>
    public static readonly PropertyDefinition Flag = new ComplexPropertyDefinition<Flag>(
        XmlElementNames.Flag,
        FieldUris.Flag,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        () => new Flag()
    );

    /// <summary>
    ///     Defines the PolicyTag property.
    /// </summary>
    public static readonly PropertyDefinition PolicyTag = new ComplexPropertyDefinition<PolicyTag>(
        XmlElementNames.PolicyTag,
        FieldUris.PolicyTag,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        () => new PolicyTag()
    );

    /// <summary>
    ///     Defines the ArchiveTag property.
    /// </summary>
    public static readonly PropertyDefinition ArchiveTag = new ComplexPropertyDefinition<ArchiveTag>(
        XmlElementNames.ArchiveTag,
        FieldUris.ArchiveTag,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        () => new ArchiveTag()
    );

    /// <summary>
    ///     Defines the RetentionDate property.
    /// </summary>
    public static readonly PropertyDefinition RetentionDate = new DateTimePropertyDefinition(
        XmlElementNames.RetentionDate,
        FieldUris.RetentionDate,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        true
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
    ///     Defines the TextBody property.
    /// </summary>
    public static readonly PropertyDefinition TextBody = new ComplexPropertyDefinition<TextBody>(
        XmlElementNames.TextBody,
        FieldUris.TextBody,
        PropertyDefinitionFlags.MustBeExplicitlyLoaded,
        ExchangeVersion.Exchange2013,
        () => new TextBody()
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
    ///     Defines the Hashtags property.
    /// </summary>
    public static readonly PropertyDefinition Hashtags = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.Hashtags,
        FieldUris.Hashtags,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete,
        ExchangeVersion.Exchange2015,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the Mentions property.
    /// </summary>
    public static readonly PropertyDefinition Mentions = new ComplexPropertyDefinition<EmailAddressCollection>(
        XmlElementNames.Mentions,
        FieldUris.Mentions,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete,
        ExchangeVersion.Exchange2015,
        () => new EmailAddressCollection()
    );

    /// <summary>
    ///     Defines the MentionedMe property.
    /// </summary>
    public static readonly PropertyDefinition MentionedMe = new BoolPropertyDefinition(
        XmlElementNames.MentionedMe,
        FieldUris.MentionedMe,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2015,
        true
    );

    // This must be declared after the property definitions
    internal static readonly ItemSchema Instance = new();

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

        RegisterProperty(MimeContent);
        RegisterProperty(Id);
        RegisterProperty(ParentFolderId);
        RegisterProperty(ItemClass);
        RegisterProperty(Subject);
        RegisterProperty(Sensitivity);
        RegisterProperty(Body);
        RegisterProperty(Attachments);
        RegisterProperty(DateTimeReceived);
        RegisterProperty(Size);
        RegisterProperty(Categories);
        RegisterProperty(Importance);
        RegisterProperty(InReplyTo);
        RegisterProperty(IsSubmitted);
        RegisterProperty(IsDraft);
        RegisterProperty(IsFromMe);
        RegisterProperty(IsResend);
        RegisterProperty(IsUnmodified);
        RegisterProperty(InternetMessageHeaders);
        RegisterProperty(DateTimeSent);
        RegisterProperty(DateTimeCreated);
        RegisterProperty(AllowedResponseActions);
        RegisterProperty(ReminderDueBy);
        RegisterProperty(IsReminderSet);
        RegisterProperty(ReminderMinutesBeforeStart);
        RegisterProperty(DisplayCc);
        RegisterProperty(DisplayTo);
        RegisterProperty(HasAttachments);
        RegisterProperty(ExtendedProperties);
        RegisterProperty(Culture);
        RegisterProperty(EffectiveRights);
        RegisterProperty(LastModifiedName);
        RegisterProperty(LastModifiedTime);
        RegisterProperty(IsAssociated);
        RegisterProperty(WebClientReadFormQueryString);
        RegisterProperty(WebClientEditFormQueryString);
        RegisterProperty(ConversationId);
        RegisterProperty(UniqueBody);
        RegisterProperty(Flag);
        RegisterProperty(StoreEntryId);
        RegisterProperty(InstanceKey);
        RegisterProperty(NormalizedBody);
        RegisterProperty(EntityExtractionResult);
        RegisterProperty(PolicyTag);
        RegisterProperty(ArchiveTag);
        RegisterProperty(RetentionDate);
        RegisterProperty(Preview);
        RegisterProperty(TextBody);
        RegisterProperty(IconIndex);
        RegisterProperty(MimeContentUTF8);

        RegisterProperty(Hashtags);
        RegisterProperty(Mentions);
        RegisterProperty(MentionedMe);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemSchema" /> class.
    /// </summary>
    internal ItemSchema()
    {
    }
}
