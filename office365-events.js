Office365 = {};

Office365.event = {
    runByUser : function(userOrUserId){
        if (!userOrUserId) {
            throwError("Missing argument", "userOrUserId");
        }

        var user = null;
        if (_.isString(userOrUserId)) {
            user = Meteor.users.findOne(userOrUserId);
        }
        else if (testIsUserObject(userOrUserId)){
            user = userOrUserId;
        }
        else{
            throwError("Argument incorrect type", "userIdsOrUsers", "The supplied argument was not a user or user id");
        }

        assertUserWithAzureAdServiceDefined(user);

        return buildFluentObject(user);
    }
}
Office365.event.baseUrl = "https://outlook.office365.com/api/v1.0/me/events"

function buildFluentObject(user){
    var restRequest = {};

    var fluentObject = {
        subject : function(subject){
            restRequest.Subject = subject;
            return fluentObject;
        },
        bodyHtml : function(content){
            restRequest.Body = {
                ContentType : "HTML",
                Content : content
            };
            return fluentObject;
        },
        bodyPlainText : function(content){
            restRequest.Body = {
                ContentType : "Text",
                Content : content
            };
            return fluentObject;
        },
        attendees : function(userIdsOrUsers){
            var users;

            //check the type of the parameter and extract the users
            if (_.isString(userIdsOrUsers)) {
                users = [Meteor.users.findOne(userIdsOrUsers)];
            }
            else if (testIsUserObject(userIdsOrUsers)){
                users = [userIdsOrUsers];
            }
            else if (_.isArray(userIdsOrUsers) && _.isEmpty(userIdsOrUsers)){
                throwError("Argument list is empty", "userIdsOrUsers");
            }
            else if (_.isArray(userIdsOrUsers) && _.isString(_.first(userIdsOrUsers))){
                users = Meteor.users.find({ _id : { $in : userIdsOrUsers}}).fetch();
            }
            else if (_.isArray(userIdsOrUsers) && testIsUserObject(_.first(userIdsOrUsers))){
                users = userIdsOrUsers;
            }
            else{
                throwError(
                    "Unexpected type of argument",
                    "userIdsOrUsers",
                    "userIdsOrUsers must be either: 1.) a user id 2.) a user 3.) a list of user ids 4.) a list of users"
                );
            }

            _.each(users, assertUserWithAzureAdServiceDefined);

            restRequest.Attendees = _.map(users, function(user){
                return {
                    EmailAddress : {
                        Name : user.services.azureAd.displayName,
                        Address : user.services.azureAd.mail || user.services.azureAd.userPrincipleName
                    }
                }
            });

            return fluentObject;
        },
        location : function(name){
            restRequest.Location = {
                DisplayName : name
            };

            return fluentObject;
        },
        startsAt : function(momentTz){
            checkForMomentTimeZone(momentTz);
            restRequest.Start = momentTz.format();
            return fluentObject;
        },
        endsAt : function(momentTz){
            checkForMomentTimeZone(momentTz);
            restRequest.End = momentTz.format();
            return fluentObject;
        },
        requireAResponse : function(responseRequested){
            restRequest.ResponseRequested = responseRequested;
            return fluentObject;
        },
        update : function(eventId){
            return callOffice365ApiMethod(user, "PATCH", getEventSpecificUrl(eventId), { data : restRequest});
        },
        create : function(){
            return callOffice365ApiMethod(user, "POST", Office365.event.baseUrl, { data : restRequest});
        },
        delete : function(eventId){
            callOffice365ApiMethod(user, "DELETE", getEventSpecificUrl(eventId));
        },
        get : function(eventId) {
            return callOffice365ApiMethod(user, "GET", getEventSpecificUrl(eventId));
        }
    };

    return fluentObject;
}

function getEventSpecificUrl(eventId){
    return Office365.event.baseUrl  + "/" + eventId;
}

function callOffice365ApiMethod(user, method, url, restRequest){
    var accessToken = AzureAd.resources.getOrUpdateUserAccessToken(AzureAd.resources.office365.friendlyName, user);
    var requestOptions;
    if (restRequest) {
        requestOptions = {data: restRequest};
    }

    return AzureAd.http.callAuthenticated(method, url, accessToken, restRequest);
}

function checkForMomentTimeZone(obj){
    if (!obj || !obj._isAMomentObject || !obj._z)
        throw new Meteor.Error("office365-events:Invalid argument type", "momentTz", "Please supply a moment timezone");
}

function testIsUserObject(obj){
    return Match.test(obj, Match.ObjectIncluding({
        _id : String
    }));
}

function assertUserWithAzureAdServiceDefined(user){
    var isValid = "services" in user && "azureAd" in user.services && _.isObject(user.services.azureAd);

    if (!isValid) {
        throwError(
            "User is missing AzureAd credentials",
            "The user has not authenticated yet.",
            "Ensure the user has authenticated via the accounts-azure-active-directory/azure-active-directory package"
        );
    }
}

function throwError(error, reason, details){
    throw new Meteor.Error("office365-events:" + error, reason, details);
};