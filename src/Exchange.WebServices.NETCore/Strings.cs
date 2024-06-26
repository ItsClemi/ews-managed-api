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

namespace Microsoft.Exchange.WebServices;

internal static class Strings
{
    internal const string CannotRemoveSubscriptionFromLiveConnection =
        "Subscriptions can't be removed from an open connection.";

    internal const string ReadAccessInvalidForNonCalendarFolder =
        "The Permission read access value {0} can't be used with a non-calendar folder.";

    internal const string PropertyDefinitionPropertyMustBeSet = "The PropertyDefinition property must be set.";

    internal const string ArgumentIsBlankString = "The string argument contains only white space characters.";
    internal const string InvalidAutodiscoverDomainsCount = "At least one domain name must be requested.";
    internal const string MinutesMustBeBetween0And1439 = "minutes must be between 0 and 1439, inclusive.";

    internal const string DeleteInvalidForUnsavedUserConfiguration =
        "This user configuration object can't be deleted because it's never been saved.";

    internal const string PeriodNotFound = "Invalid transition. A period with the specified Id couldn't be found: {0}";

    internal const string InvalidAutodiscoverSmtpAddress = "A valid SMTP address must be specified.";
    internal const string InvalidOAuthToken = "The given token is invalid.";
    internal const string MaxScpHopsExceeded = "The number of SCP URL hops exceeded the limit.";

    internal const string ContactGroupMemberCannotBeUpdatedWithoutBeingLoadedFirst =
        "The contact group's Members property must be reloaded before newly-added members can be updated.";

    internal const string CurrentPositionNotElementStart = "The current position is not the start of an element.";

    internal const string CannotConvertBetweenTimeZones = "Unable to convert {0} from {1} to {2}.";

    internal const string FrequencyMustBeBetween1And1440 = "The frequency must be a value between 1 and 1440.";

    internal const string CannotSetDelegateFolderPermissionLevelToCustom =
        "This operation can't be performed because one or more folder permission levels were set to Custom.";

    internal const string PartnerTokenIncompatibleWithRequestVersion =
        "TryGetPartnerAccess only supports {0} or a later version in Microsoft-hosted data center.";

    internal const string InvalidAutodiscoverRequest = "Invalid Autodiscover request: '{0}'";

    internal const string InvalidAsyncResult =
        "The IAsyncResult object was not returned from the corresponding asynchronous method of the original ExchangeService object.";

    internal const string InvalidMailboxType = "The mailbox type isn't valid.";
    internal const string AttachmentCollectionNotLoaded = "The attachment collection must be loaded.";

    internal const string ParameterIncompatibleWithRequestVersion =
        "The parameter {0} is only valid for Exchange Server version {1} or a later version.";

    internal const string DayOfWeekIndexMustBeSpecifiedForRecurrencePattern =
        "The recurrence pattern's DayOfWeekIndex property must be specified.";

    internal const string WLIDCredentialsCannotBeUsedWithLegacyAutodiscover =
        "This type of credentials can't be used with this AutodiscoverService.";

    internal const string PropertyCannotBeUpdated = "This property can't be updated.";
    internal const string IncompatibleTypeForArray = "Type {0} can't be used as an array of type {1}.";

    internal const string AutodiscoverServiceIncompatibleWithRequestVersion =
        "The Autodiscover service only supports {0} or a later version.";

    internal const string InvalidAutodiscoverSmtpAddressesCount = "At least one SMTP address must be requested.";

    internal const string ServiceUrlMustBeSet = "The Url property on the ExchangeService object must be set.";

    internal const string ItemTypeNotCompatible =
        "The item type returned by the service ({0}) isn't compatible with the requested item type ({1}).";

    internal const string AttachmentItemTypeMismatch =
        "Can not update this attachment item since the item in the response has a different type.";

    internal const string UnsupportedWebProtocol = "Protocol {0} isn't supported for service requests.";

    internal const string EnumValueIncompatibleWithRequestVersion =
        "Enumeration value {0} in enumeration type {1} is only valid for Exchange version {2} or later.";

    internal const string UnexpectedElement =
        "An element node '{0}:{1}' of the type {2} was expected, but node '{3}' of type {4} was found.";

    internal const string NoAppropriateConstructorForItemClass =
        "No appropriate constructor could be found for this item class.";

    internal const string SearchFilterAtIndexIsInvalid = "The search filter at index {0} is invalid.";

    internal const string DeletingThisObjectTypeNotAuthorized = "Deleting this type of object isn't authorized.";

    internal const string PropertyCannotBeDeleted = "This property can't be deleted.";
    internal const string ValuePropertyMustBeSet = "The Value property must be set.";

    internal const string TagValueIsOutOfRange = "The extended property tag value must be in the range of 0 to 65,535.";

    internal const string ItemToUpdateCannotBeNullOrNew = "Items[{0}] is either null or does not have an Id.";

    internal const string SearchParametersRootFolderIdsEmpty = "SearchParameters must contain at least one folder id.";

    internal const string MailboxQueriesParameterIsNotSpecified =
        "The collection of query and mailboxes parameter is not specified.";

    internal const string FolderPermissionHasInvalidUserId =
        "The UserId in the folder permission at index {0} is invalid. The StandardUser, PrimarySmtpAddress, or SID property must be set.";

    internal const string InvalidAutodiscoverDomain = "The domain name must be specified.";

    internal const string MailboxesParameterIsNotSpecified = "The array of mailboxes (in legacy DN) is not specified.";

    internal const string DayOfMonthMustBeSpecifiedForRecurrencePattern =
        "The recurrence pattern's DayOfMonth property must be specified.";

    internal const string ClassIncompatibleWithRequestVersion =
        "Class {0} is only valid for Exchange version {1} or later.";

    internal const string CertificateHasNoPrivateKey =
        "The given certificate does not have the private key. The private key is necessary to sign part of the request message.";

    internal const string InvalidOrUnsupportedTimeZoneDefinition =
        "The time zone definition is invalid or unsupported.";

    internal const string HourMustBeBetween0And23 = "Hour must be between 0 and 23.";
    internal const string TimeoutMustBeBetween1And1440 = "Timeout must be a value between 1 and 1440.";
    internal const string CredentialsRequired = "Credentials are required to make a service request.";

    internal const string MustLoadOrAssignPropertyBeforeAccess =
        "You must load or assign this property before you can read its value.";

    internal const string InvalidAutodiscoverServiceResponse = "The Autodiscover service response was invalid.";

    internal const string CannotCallConnectDuringLiveConnection = "The connection has already opened.";
    internal const string ObjectDoesNotHaveId = "This service object doesn't have an ID.";

    internal const string CannotAddSubscriptionToLiveConnection = "Subscriptions can't be added to an open connection.";

    internal const string MaxChangesMustBeBetween1And512 = "MaxChangesReturned must be between 1 and 512.";

    internal const string AttributeValueCannotBeSerialized =
        "Values of type '{0}' can't be used for the '{1}' attribute.";

    internal const string NumberOfDaysMustBePositive = "NumberOfDays must be zero or greater. Zero indicates no limit.";

    internal const string SearchFilterMustBeSet = "The SearchFilter property must be set.";
    internal const string EndDateMustBeGreaterThanStartDate = "EndDate must be greater than StartDate.";
    internal const string InvalidDateTime = "Invalid date and time: {0}.";

    internal const string UpdateItemsDoesNotAllowAttachments =
        "This operation can't be performed because attachments have been added or deleted for one or more items.";

    internal const string TimeoutMustBeGreaterThanZero = "Timeout must be greater than zero.";

    internal const string AutodiscoverInvalidSettingForOutlookProvider =
        "The requested setting, '{0}', isn't supported by this Autodiscover endpoint.";

    internal const string InvalidRedirectionResponseReturned = "The service returned an invalid redirection response.";

    internal const string ExpectedStartElement =
        "The start element was expected, but node '{0}' of type {1} was found.";

    internal const string DaysOfTheWeekNotSpecified =
        "The recurrence pattern's property DaysOfTheWeek must contain at least one day of the week.";

    internal const string FolderToUpdateCannotBeNullOrNew = "Folders[{0}] is either null or does not have an Id.";

    internal const string PartnerTokenRequestRequiresUrl =
        "TryGetPartnerAccess request requires the Url be set with the partner's autodiscover url first.";

    internal const string NumberOfOccurrencesMustBeGreaterThanZero = "NumberOfOccurrences must be greater than 0.";

    internal const string StartTimeZoneRequired =
        "StartTimeZone required when setting the Start, End, IsAllDayEvent, or Recurrence properties.  You must load or assign this property before attempting to update the appointment.";

    internal const string PropertyAlreadyExistsInOrderByCollection =
        "Property {0} already exists in OrderByCollection.";

    internal const string ItemAttachmentMustBeNamed = "The name of the item attachment at index {0} must be set.";

    internal const string InvalidAutodiscoverSettingsCount = "At least one setting must be requested.";
    internal const string LoadingThisObjectTypeNotSupported = "Loading this type of object is not supported.";

    internal const string UserIdForDelegateUserNotSpecified = "The UserId in the DelegateUser hasn't been specified.";

    internal const string PhoneCallAlreadyDisconnected = "The phone call has already been disconnected.";

    internal const string OperationDoesNotSupportAttachments = "This operation isn't supported on attachments.";

    internal const string UnsupportedTimeZonePeriodTransitionTarget =
        "The time zone transition target isn't supported.";

    internal const string IEnumerableDoesNotContainThatManyObject =
        "The IEnumerable doesn't contain that many objects.";

    internal const string UpdateItemsDoesNotSupportNewOrUnchangedItems =
        "This operation can't be performed because one or more items are new or unmodified.";

    internal const string ValidationFailed = "Validation failed.";
    internal const string InvalidRecurrencePattern = "Invalid recurrence pattern: ({0}).";

    internal const string TimeWindowStartTimeMustBeGreaterThanEndTime =
        "The time window's end time must be greater than its start time.";

    internal const string InvalidAttributeValue = "The invalid value '{0}' was specified for the '{1}' attribute.";

    internal const string FileAttachmentContentIsNotSet =
        "The content of the file attachment at index {0} must be set.";

    internal const string AutodiscoverDidNotReturnEwsUrl =
        "The Autodiscover service didn't return an appropriate URL that can be used for the ExchangeService Autodiscover URL.";

    internal const string RecurrencePatternMustHaveStartDate =
        "The recurrence pattern's StartDate property must be specified.";

    internal const string OccurrenceIndexMustBeGreaterThanZero = "OccurrenceIndex must be greater than 0.";

    internal const string ServiceResponseDoesNotContainXml =
        "The response received from the service didn't contain valid XML.";

    internal const string ItemIsOutOfDate =
        "The operation can't be performed because the item is out of date. Reload the item and try again.";

    internal const string MinuteMustBeBetween0And59 = "Minute must be between 0 and 59.";

    internal const string NoSoapOrWsSecurityEndpointAvailable =
        "No appropriate Autodiscover SOAP or WS-Security endpoint is available.";

    internal const string ElementNotFound =
        "The element '{0}' in namespace '{1}' wasn't found at the current position.";

    internal const string IndexIsOutOfRange = "index is out of range.";
    internal const string PropertyIsReadOnly = "This property is read-only and can't be set.";
    internal const string AttachmentCreationFailed = "At least one attachment couldn't be created.";
    internal const string DayOfMonthMustBeBetween1And31 = "DayOfMonth must be between 1 and 31.";
    internal const string ServiceRequestFailed = "The request failed. {0}";

    internal const string DelegateUserHasInvalidUserId =
        "The UserId in the DelegateUser is invalid. The StandardUser, PrimarySmtpAddress or SID property must be set.";

    internal const string SearchFilterComparisonValueTypeIsNotSupported =
        "Values of type '{0}' can't be used as comparison values in search filters.";

    internal const string ElementValueCannotBeSerialized = "Values of type '{0}' can't be used for the '{1}' element.";

    internal const string PropertyValueMustBeSpecifiedForRecurrencePattern =
        "The recurrence pattern's {0} property must be specified.";

    internal const string NonSummaryPropertyCannotBeUsed = "The property {0} can't be used in {1} requests.";
    internal const string HoldIdParameterIsNotSpecified = "The hold id parameter is not specified.";

    internal const string TransitionGroupNotFound =
        "Invalid transition. A transition group with the specified ID couldn't be found: {0}";

    internal const string ObjectTypeNotSupported =
        "Objects of type {0} can't be added to the dictionary. The following types are supported: string array, byte array, boolean, byte, DateTime, integer, long, string, unsigned integer, and unsigned long.";

    internal const string InvalidTimeoutValue = "{0} is not a valid timeout value. Valid values range from 1 to 1440.";

    internal const string AutodiscoverRedirectBlocked =
        "Autodiscover blocked a potentially insecure redirection to {0}. To allow Autodiscover to follow the redirection, use the AutodiscoverUrl(string, AutodiscoverRedirectionUrlValidationCallback) overload.";

    internal const string PropertySetCannotBeModified = "This PropertySet is read-only and can't be modified.";

    internal const string DayOfTheWeekMustBeSpecifiedForRecurrencePattern =
        "The recurrence pattern's property DayOfTheWeek must be specified.";

    internal const string ServiceObjectAlreadyHasId =
        "This operation can't be performed because this service object already has an ID. To update this service object, use the Update() method instead.";

    internal const string MethodIncompatibleWithRequestVersion =
        "Method {0} is only valid for Exchange Server version {1} or later.";

    internal const string OperationNotSupportedForPropertyDefinitionType =
        "This operation isn't supported for property definition type {0}.";

    internal const string InvalidElementStringValue = "The invalid value '{0}' was specified for the '{1}' element.";

    internal const string CollectionIsEmpty = "The collection is empty.";

    internal const string InvalidFrequencyValue =
        "{0} is not a valid frequency value. Valid values range from 1 to 1440.";

    internal const string UnexpectedEndOfXmlDocument = "The XML document ended unexpectedly.";

    internal const string FolderTypeNotCompatible =
        "The folder type returned by the service ({0}) isn't compatible with the requested folder type ({1}).";

    internal const string RequestIncompatibleWithRequestVersion =
        "The service request {0} is only valid for Exchange version {1} or later.";

    internal const string PropertyTypeIncompatibleWhenUpdatingCollection =
        "Can not update the existing collection item since the item in the response has a different type.";

    internal const string ServerVersionNotSupported = "Exchange Server doesn't support the requested version.";

    internal const string DurationMustBeSpecifiedWhenScheduled =
        "Duration must be specified when State is equal to Scheduled.";

    internal const string NoError = "No error.";

    internal const string CannotUpdateNewUserConfiguration =
        "This user configuration can't be updated because it's never been saved.";

    internal const string ObjectTypeIncompatibleWithRequestVersion =
        "The object type {0} is only valid for Exchange Server version {1} or later versions.";

    internal const string NullStringArrayElementInvalid = "The array contains at least one null element.";
    internal const string HttpsIsRequired = "Https is required when partner token is expected.";

    internal const string MergedFreeBusyIntervalMustBeSmallerThanTimeWindow =
        "MergedFreeBusyInterval must be smaller than the specified time window.";

    internal const string SecondMustBeBetween0And59 = "Second must be between 0 and 59.";

    internal const string AtLeastOneAttachmentCouldNotBeDeleted = "At least one attachment couldn't be deleted.";

    internal const string IdAlreadyInList = "The ID is already in the list.";

    internal const string BothSearchFilterAndQueryStringCannotBeSpecified =
        "Both search filter and query string can't be specified. One of them must be null.";

    internal const string AdditionalPropertyIsNull = "The additional property at index {0} is null.";
    internal const string InvalidEmailAddress = "The e-mail address is formed incorrectly.";

    internal const string MaximumRedirectionHopsExceeded = "The maximum redirection hop count has been reached.";

    internal const string AutodiscoverCouldNotBeLocated = "The Autodiscover service couldn't be located.";

    internal const string NoSubscriptionsOnConnection =
        "You must add at least one subscription to this connection before it can be opened.";

    internal const string PermissionLevelInvalidForNonCalendarFolder =
        "The Permission level value {0} can't be used with a non-calendar folder.";

    internal const string InvalidAuthScheme = "The token auth scheme should be bearer.";

    internal const string ValuePropertyNotLoaded = "This property was requested, but it wasn't returned by the server.";

    internal const string PropertyIncompatibleWithRequestVersion =
        "The property {0} is valid only for Exchange {1} or later versions.";

    internal const string OffsetMustBeGreaterThanZero = "The offset must be greater than 0.";

    internal const string CreateItemsDoesNotAllowAttachments =
        "This operation doesn't support items that have attachments.";

    internal const string PropertyDefinitionTypeMismatch =
        "Property definition type '{0}' and type parameter '{1}' aren't compatible.";

    internal const string IntervalMustBeGreaterOrEqualToOne = "The interval must be greater than or equal to 1.";

    internal const string CannotSetPermissionLevelToCustom =
        "The PermissionLevel property can't be set to FolderPermissionLevel.Custom. To define a custom permission, set its individual properties to the values you want.";

    internal const string ArrayMustHaveAtLeastOneElement = "The Array value must have at least one element.";

    internal const string MonthMustBeSpecifiedForRecurrencePattern =
        "The recurrence pattern's Month property must be specified.";

    internal const string ValueOfTypeCannotBeConverted =
        "The value '{0}' of type {1} can't be converted to a value of type {2}.";

    internal const string ValueCannotBeConverted = "The value '{0}' couldn't be converted to type {1}.";
    internal const string ServerErrorAndStackTraceDetails = "{0} -- Server Error: {1}: {2} {3}";

    internal const string AutodiscoverError = "The Autodiscover service returned an error.";
    internal const string ArrayMustHaveSingleDimension = "The array value must have a single dimension.";
    internal const string InvalidPropertyValueNotInRange = "{0} must be between {1} and {2}.";

    internal const string RegenerationPatternsOnlyValidForTasks =
        "Regeneration patterns can only be used with Task items.";

    internal const string ItemAttachmentCannotBeUpdated = "Item attachments can't be updated.";

    internal const string EqualityComparisonFilterIsInvalid =
        "Either the OtherPropertyDefinition or the Value properties must be set.";

    internal const string AutodiscoverServiceRequestRequiresDomainOrUrl =
        "This Autodiscover request requires that either the Domain or Url be specified.";

    internal const string InvalidUser = "Invalid user: '{0}'";
    internal const string AccountIsLocked = "This account is locked. Visit {0} to unlock it.";
    internal const string InvalidDomainName = "'{0}' is not a valid domain name.";

    internal const string TooFewServiceReponsesReturned =
        "The service was expected to return {1} responses of type '{0}', but {2} responses were received.";

    internal const string CannotSubscribeToStatusEvents = "Status events can't be subscribed to.";

    internal const string InvalidSortByPropertyForMailboxSearch = "Specified SortBy property '{0}' is invalid.";

    internal const string UnexpectedElementType = "The expected XML node type was {0}, but the actual type is {1}.";

    internal const string ValueMustBeGreaterThanZero = "The value must be greater than 0.";
    internal const string AttachmentCannotBeUpdated = "Attachments can't be updated.";

    internal const string CreateItemsDoesNotHandleExistingItems =
        "This operation can't be performed because at least one item already has an ID.";

    internal const string MultipleContactPhotosInAttachment =
        "This operation only allows at most 1 file attachment with IsContactPhoto set.";

    internal const string InvalidRecurrenceRange = "Invalid recurrence range: ({0}).";

    internal const string CannotSetBothImpersonatedAndPrivilegedUser =
        "Can't set both impersonated user and privileged user in the ExchangeService object.";

    internal const string CannotCallDisconnectWithNoLiveConnection = "The connection is already closed.";
    internal const string IdPropertyMustBeSet = "The Id property must be set.";

    internal const string ValuePropertyNotAssigned = "You must assign this property before you can read its value.";

    internal const string ZeroLengthArrayInvalid = "The array must contain at least one element.";

    internal const string HoldMailboxesParameterIsNotSpecified = "The hold mailboxes parameter is not specified.";

    internal const string CannotSaveNotNewUserConfiguration =
        "Calling Save isn't allowed because this user configuration isn't new. To apply local changes to this user configuration, call Update instead.";

    internal const string ServiceObjectDoesNotHaveId =
        "This operation can't be performed because this service object doesn't have an Id.";

    internal const string XsDurationCouldNotBeParsed = "The specified xsDuration argument couldn't be parsed.";

    internal const string UnknownTimeZonePeriodTransitionType = "Unknown time zone transition type: {0}";
    internal const string UserPhotoSizeNotSpecified = "The UserPhotoSize must be not be null or empty.";
    internal const string UserPhotoNotSpecified = "The photo must be not be null or empty.";
    internal const string PercentCompleteMustBeBetween0And100 = "PercentComplete must be between 0 and 100.";

    internal const string InvalidOrderBy = "At least one of the property definitions in the OrderBy clause is null.";

    internal const string ParentFolderDoesNotHaveId = "parentFolder doesn't have an Id.";

    internal const string CannotAddRequestHeader =
        "HTTP header '{0}' isn't permitted. Only HTTP headers with the 'X-' prefix are permitted.";

    internal const string FolderPermissionLevelMustBeSet =
        "The permission level of the folder permission at index {0} must be set.";

    internal const string NewMessagesWithAttachmentsCannotBeSentDirectly =
        "New messages with attachments can't be sent directly. You must first save the message and then send it.";

    internal const string PropertyCollectionSizeMismatch =
        "The collection returned by the service has a different size from the current one.";
}
