#Office-365-events

A meteor package that exposes a fluent interface for working with the office 365 events REST API.

##How do I use this package?
Please refer to the [Wiki](https://github.com/djluck/office365-events/wiki)

##Sample code
    var createdEvent = Office365.event
        .runByUser(userId)
        .subject("This is the subject")
        .bodyHtml("<b>This is the body</b> Its using html!")
        .attendees(Meteor.users.find().fetch())
        .location("The test room")
        .startsAt(moment().add(2, "days").tz("Europe/London"))
        .endsAt(moment().add(2, "days").add(1, "hour").tz("Europe/London"))
        .requireAResponse(false)
        .create();

The [Wiki](https://github.com/djluck/office365-events/wiki) contains more examples and notes on using this package.