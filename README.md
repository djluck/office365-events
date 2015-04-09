#Office-365-events

A meteor package that exposes a fluent interface for working with the office 365 events REST API.

##Authenticating
This package includes the [azure-resource-office-365](https://atmospherejs.com/wiseguyeh/azure-resource-office-365) package which is used to authenticate against Azure Active Directory.
Please include the [accounts-azure-active-directory](https://atmospherejs.com/wiseguyeh/accounts-azure-active-directory) package if you wish to allow your users to log in using the `{{> loginButtons}}` template.

##Examples

###Create an event

    var createdEvent = Office365.event
        //events are owned by a user. runByUser accepts a user id or user.
        .runByUser(userId)
        .subject("This is the subject")
        .bodyHtml("<b>This is the body</b> Its using html!")
        //attendees accepts a user id, an array of user ids, a user or an array of users
        .attendees(Meteor.users.find())
        .location("The test room")
        //Office 365 requires a UTC time for start/end date times.
        //startsAt/endsAt require a moment timezone object to specify the correct UTC time and zone offset
        .startsAt(moment().add(2, "days").tz("Europe/London"))
        .endsAt(moment().add(2, "days").add(1, "hour").tz("Europe/London"))
        .requireAResponse(false)
        .create();

###Update an event
    //only included fields will be updated. If you don't change a field, then it won't be updated.
    var updatedEvent = Office365.event
        .runByUser(user)
        .subject("This is the new subject")
        //bodies can be either HTML or plain text
        .bodyPlainText("This is the new plain text body")
        //when updating attendees you cannot add or remove attendees, only replace the entire collection of attendees.
        .attendees([userId1, userId2])
        .requireAResponse(true)
        .update(eventId);

###Get an existing event
    var event = Office365.event
        .runByUser(user)
        .get(eventId)

###Delete an event
    Office365.event
        .runByUser(user)
        .delete(eventId)


##Notes on use
 - `create`, `update`, and `get` all return an [Event object](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#EventResource)
 - Event objects contain the property `Id` which you can use to refer to an event when calling `update`, `delete` or `get`.
 - Event id's are not stored in any way. This is left up to the consumer of the package.
 - Any user that owns an event or is attending an event MUST have been authenticated against Azure Active Directory.





