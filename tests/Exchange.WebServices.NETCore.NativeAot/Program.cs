using System.Reflection;

using Microsoft.Exchange.WebServices.Data;

// Literal type/method names keep these internal entry points statically analyzable without
// rooting every EWS field in the test and accidentally hiding missing library annotations.
var serialize =
    Type.GetType("Microsoft.Exchange.WebServices.Data.EwsUtilities, Exchange.WebServices.NETCore", true)!.GetMethod(
        "SerializeEnum",
        BindingFlags.Static | BindingFlags.NonPublic
    )!.CreateDelegate<Func<Enum, string>>();
var validateVersion =
    Type.GetType("Microsoft.Exchange.WebServices.Data.EwsUtilities, Exchange.WebServices.NETCore", true)!.GetMethod(
        "ValidateEnumVersionValue",
        BindingFlags.Static | BindingFlags.NonPublic
    )!.CreateDelegate<Action<Enum, ExchangeVersion>>();
var initializeNames =
    typeof(ServiceObjectSchema).GetMethod(
        "InitializeSchemaPropertyNames",
        BindingFlags.Static | BindingFlags.NonPublic
    )!.CreateDelegate<Action>();
var findProperty =
    typeof(ServiceObjectSchema).GetMethod("FindPropertyDefinition", BindingFlags.Static | BindingFlags.NonPublic)!
        .CreateDelegate<Func<string, PropertyDefinitionBase>>();

Check(serialize(EventType.NewMail) == "NewMailEvent", "Event schema name");
Check(serialize(WellKnownFolderName.Inbox) == "inbox", "Folder schema name");
Check(serialize(MailboxType.ContactGroup) == "PrivateDL", "Mailbox schema name");
Check(serialize(ItemTraversal.Shallow) == "Shallow", "Unmapped enum name");

validateVersion(WellKnownFolderName.Inbox, ExchangeVersion.Exchange2007_SP1);
validateVersion(EventType.FreeBusyChanged, ExchangeVersion.Exchange2010_SP1);
try
{
    validateVersion(EventType.FreeBusyChanged, ExchangeVersion.Exchange2007_SP1);
    throw new InvalidOperationException("The enum version attribute was lost.");
}
catch (ServiceVersionException)
{
    // The field attribute must survive linking to reject older servers.
}

initializeNames();
Check(findProperty("item:Importance") is PropertyDefinition { Name: "Importance" }, "Unreferenced field metadata");
Check(ReferenceEquals(findProperty("item:Subject"), ItemSchema.Subject), "Item schema lookup");
Check(ReferenceEquals(findProperty("calendar:Start"), AppointmentSchema.Start), "Calendar schema lookup");
Check(ReferenceEquals(findProperty("contacts:FileAs"), ContactSchema.FileAs), "Contact schema lookup");
Check(ReferenceEquals(findProperty("task:DueDate"), TaskSchema.DueDate), "Task schema lookup");
Check(ReferenceEquals(findProperty("folder:DisplayName"), FolderSchema.DisplayName), "Folder schema lookup");
Console.WriteLine("EWS metadata smoke checks passed.");
return;

static void Check(bool condition, string message)
{
    if (!condition)
    {
        throw new InvalidOperationException(message);
    }
}
