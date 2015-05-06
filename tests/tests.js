Tinytest.add("runByUser() accepts a user id", function(){
    Office365.event.runByUser(arrangeUserWithAzureCredentials()._id);
});

Tinytest.add("runByUser() accepts a user object", function(){
    Office365.event.runByUser(arrangeUserWithAzureCredentials());
});

Tinytest.add("runByUser() rejects a user object without the azureAd service", function(test){
    test.throwsException(
        Office365.event.runByUser.bind(this, arrangeUser()),
        "User is missing AzureAd credentials"
    );
});

Tinytest.add("runByUser() rejects a user id for a user object without the azureAd service", function(test){
    test.throwsException(
        Office365.event.runByUser.bind(this, arrangeUser()._id),
        "User is missing AzureAd credentials"
    );
});

Tinytest.add("runByUser() rejects any other value", function(test){
    test.throwsException(
        Office365.event.runByUser.bind(this, 1),
        "Argument incorrect type"
    );
    test.throwsException(
        Office365.event.runByUser.bind(this, null),
        "Argument incorrect type"
    );
    test.throwsException(
        Office365.event.runByUser.bind(this, {}),
        "Argument incorrect type"
    );
});



Tinytest.addWithSinon("attendees accepts a user id and sets the attendees of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user = arrangeUserWithAzureCredentials("test1@email.com", "My Name");

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).attendees(user._id),
        eventRunner,
        {
            Attendees : [
                {
                    EmailAddress : {
                        Address : "test1@email.com",
                        Name : "My Name"
                    }
                }
            ]
        }
    )
});

Tinytest.addWithSinon("attendees rejects a user id for a user object without the azureAd service", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user = arrangeUser("test1@email.com", "My Name");

    //Act + Assert:
    test.throwsException(
        Office365.event.runByUser(eventRunner).attendees.bind(this, user._id),
        "User is missing AzureAd credentials"
    )
});

Tinytest.addWithSinon("attendees accepts a user object and sets the attendees of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user = arrangeUserWithAzureCredentials("test2@email.com", "My Name 2");

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).attendees(user),
        eventRunner,
        {
            Attendees : [
                {
                    EmailAddress : {
                        Address : "test2@email.com",
                        Name : "My Name 2"
                    }
                }
            ]
        }
    )
});

Tinytest.addWithSinon("attendees rejects a user object without the azureAd service", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user = arrangeUser("test2@email.com", "My Name 2");

    //Act + Assert:
    test.throwsException(
        Office365.event.runByUser(eventRunner).attendees.bind(this, user._id),
        "User is missing AzureAd credentials"
    )
});

Tinytest.addWithSinon("attendees accepts an email address and sets the attendees of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).attendees("test3@email.com"),
        eventRunner,
        {
            Attendees : [
                {
                    EmailAddress : {
                        Address : "test3@email.com"
                    }
                }
            ]
        }
    )
});

Tinytest.addWithSinon("attendees accepts an array of user id's, user objects or email addresses and sets the attendees of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user1 = arrangeUserWithAzureCredentials("test1@email.com", "My Name 1");
    var user2 = arrangeUserWithAzureCredentials("test2@email.com", "My Name 2");
    var user3 = arrangeUserWithAzureCredentials("test3@email.com", "My Name 3");
    var address1 = "test4@email.com";
    var address2 = "test5@email.com";

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).attendees([address2, user3, user2._id, user1, address1]),
        eventRunner,
        {
            Attendees : [
                //although user2 was in the middle of the attendee array, because we passed a user id
                //we load the user object and put it to the front of the list
                {
                    EmailAddress : {
                        Address : "test2@email.com",
                        Name : "My Name 2"
                    }
                },
                {
                    EmailAddress : {
                        Address : "test5@email.com"
                    }
                },
                {
                    EmailAddress : {
                        Address : "test3@email.com",
                        Name : "My Name 3"
                    }
                },
                {
                    EmailAddress : {
                        Address : "test1@email.com",
                        Name : "My Name 1"
                    }
                },
                {
                    EmailAddress : {
                        Address : "test4@email.com"
                    }
                }
            ]
        }
    )
});

Tinytest.addWithSinon("attendees rejects an array if it contains a user object without the azureAd service", function(test){
//Arrange
    var eventRunner = arrangeUserWithAzureCredentials();
    var user1 = arrangeUserWithAzureCredentials("test1@email.com", "My Name 1");
    var user2 = arrangeUser("test2@email.com", "My Name 2");
    var address1 = "test4@test.com";
    var address2 = "test5@test.com";

    //Act + Assert:
    //Act + Assert:
    test.throwsException(
        Office365.event.runByUser(eventRunner).attendees.bind(this,[address1, address2, user1, user2]),
        "User is missing AzureAd credentials"
    )
});

Tinytest.addWithSinon("subject sets the subject of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).subject("A Subject"),
        eventRunner,
        {
            Subject : "A Subject"
        }
    );
})

Tinytest.addWithSinon("startsAt sets the start time of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).startsAt(moment.tz("2014-06-01 12:00", "Europe/London")),
        eventRunner,
        {
            Start : "2014-06-01T12:00:00+01:00"
        }
    );
})

Tinytest.addWithSinon("endsAt sets the finish time of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).endsAt(moment.tz("2014-01-01 12:00", "Europe/London")),
        eventRunner,
        {
            End : "2014-01-01T12:00:00+00:00"
        }
    );
});

Tinytest.addWithSinon("requireAResponse sets if a response is required from the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).requireAResponse(true),
        eventRunner,
        {
            ResponseRequested : true
        }
    );
});

Tinytest.addWithSinon("location sets the location of the event", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).location("my location"),
        eventRunner,
        {
            Location : {
                DisplayName : "my location"
            }
        }
    );
});

Tinytest.addWithSinon("bodyPlainText sets the body of the event in plain text", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).bodyPlainText("my text"),
        eventRunner,
        {
            Body : {
                ContentType : "Text",
                Content : "my text"
            }
        }
    );
});

Tinytest.addWithSinon("bodyHtml sets the body of the event in html", function(test){
    //Arrange
    var eventRunner = arrangeUserWithAzureCredentials();

    //Act + Assert:
    test.assertOffice365FluentObjectContains(
        Office365.event.runByUser(eventRunner).bodyHtml("my html"),
        eventRunner,
        {
            Body : {
                ContentType : "HTML",
                Content : "my html"
            }
        }
    );
});

Tinytest.addWithSinon("create creates the event using the office365 endpoint", function(test, sinonSandbox){
    //Arrange:
    var user = arrangeUserWithAzureCredentials("test1@email.com", "My Name 1");
    var event = Office365.event.runByUser(user)
        .subject("My Test Subject")
        .startsAt(moment.tz("2015-01-01 12:00", "America/New_York"))
        .endsAt(moment.tz("2015-01-01 13:00", "America/New_York"))
        .bodyHtml("the body")
        .attendees(user);

    var expectedApiBody = {
        Subject : "My Test Subject",
        Start : "2015-01-01T12:00:00-05:00",
        End : "2015-01-01T13:00:00-05:00",
        Attendees : [
            {
                EmailAddress : {
                    Address : "test1@email.com",
                    Name : "My Name 1"
                }
            }
        ],
        Body : {
            ContentType : "HTML",
            Content : "the body"
        }
    };
    var resultToReturn = {Id : 1};
    sinonSandbox.arrangeExpectedHttpCall(user, "https://outlook.office365.com/api/v1.0/me/events", "POST", expectedApiBody, resultToReturn);

    //Act:
    var result = event.create();

    //Assert:
    test.equal(result, resultToReturn);
});

Tinytest.addWithSinon("update updates the event using the office365 endpoint", function(test, sinonSandbox){
    //Arrange:
    var user = arrangeUserWithAzureCredentials("test1@email.com", "My Name 1");
    var event = Office365.event.runByUser(user)
        .bodyPlainText("the updated body")
        .attendees(["test2@email.com", "test3@email.com"])
        .requireAResponse(true);

    var expectedApiBody = {
        Attendees : [
            {
                EmailAddress : {
                    Address : "test2@email.com"
                }
            },
            {
                EmailAddress : {
                    Address : "test3@email.com"
                }
            }
        ],
        Body : {
            ContentType : "Text",
            Content : "the updated body"
        },
        ResponseRequested : true
    };
    var resultToReturn = {Id : 2};
    sinonSandbox.arrangeExpectedHttpCall(user, "https://outlook.office365.com/api/v1.0/me/events/2", "PATCH", expectedApiBody, resultToReturn);

    //Act:
    var result = event.update(2);

    //Assert:
    test.equal(result, resultToReturn);
});

Tinytest.addWithSinon("delete deletes the event using the office365 endpoint", function(test, sinonSandbox){
    //Arrange:
    var user = arrangeUserWithAzureCredentials("test1@email.com", "My Name 1");
    var event = Office365.event.runByUser(user);

    sinonSandbox.arrangeExpectedHttpCall(user, "https://outlook.office365.com/api/v1.0/me/events/2", "DELETE");

    //Act:
    event.delete(2);
});
//
//Tinytest.addWithSinon("get retrieves an event using the office365 endpoint", function(){
//
//});