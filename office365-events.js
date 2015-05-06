Office365 = {};

Office365.event = {
    runByUser : function(userOrUserId){
        var user = null;
        if (isObjectId(userOrUserId)) {
            user = Meteor.users.findOne(userOrUserId);
        }
        else if (isUserObject(userOrUserId)){
            user = userOrUserId;
        }
        else{
            throwError("Argument incorrect type", "userIdsOrUsers", "The supplied argument was not a user or user id", userOrUserId);
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
        attendees : function(attendees){
            /*
             Accepted argument types:
             - A single email address
             - A user id
             - A user object
             - An array, containing a mix of any of the above
             */

            restRequest.Attendees = _.isArray(attendees) ?
                getEmailAddressesFromAttendeeArray(attendees) : [getEmailAddressFromAttendee(attendees)];

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

function isUserObject(obj){
    return !!obj && _.isObject(obj) && "_id" in obj && isObjectId(obj._id);
}

function getEmailAddressesFromAttendeeArray(attendees) {
    //get all meteor ids and load them separately as it's more effecient than doing a load for each id
    var userIdsAndOthers = _.groupBy(attendees, isObjectId);
    var userIds = userIdsAndOthers.true || [];
    var others = userIdsAndOthers.false || [];
    var users = users = Meteor.users.find({ _id : { $in : userIds}}).fetch();

    return _.chain(_.union(users, others))
        .map(getEmailAddressFromAttendee)
        .value();
}

function getEmailAddressFromAttendee(attendee){
    if (isEmailAddress(attendee)){
        return {
            EmailAddress : {
                Address : attendee
            }
        };
    }
    else if (isObjectId(attendee)){
        var user = Meteor.users.findOne(attendee);
        if (!user)
            throwError("Specified user did not exist", "attendees", "You supplied an id for a user that does not exist");

        return mapUserToEmailAddress(user);
    }
    else if (isUserObject(attendee)){
        return mapUserToEmailAddress(attendee);
    }
    else{
        throwError(
            "Unexpected type of argument",
            "attendees",
            "attendees must be either: 1.) a user id 2.) a user 3.) an email address 4.) a list comprised of any of the previous 3"
        );
    }
}

function mapUserToEmailAddress(user){
    assertUserWithAzureAdServiceDefined(user);

    return {
        EmailAddress : {
            Name : user.services.azureAd.displayName,
            Address : user.services.azureAd.mail || user.services.azureAd.userPrincipleName
        }
    }
}

function isEmailAddress(obj){
    return _.isString(obj) && /.+@.+/.test(obj);
}

function isObjectId(obj){
    return _.isString(obj) && /[0-9a-zA-Z]{10,}/.test(obj);
}

function assertUserWithAzureAdServiceDefined(user){
    var isValid = "services" in user && "azureAd" in user.services && _.isObject(user.services.azureAd);

    if (!isValid) {
        throwError(
            "User is missing AzureAd credentials",
            "The user '" + user._id + "' has not authenticated yet.",
            "Ensure the user has authenticated via the accounts-azure-active-directory/azure-active-directory package"
        );
    }
}

function throwError(error, reason, details, actualValue){
    if (!!actualValue) {
        details += ". Valued passed was: " + JSON.stringify(actualValue);
    }
    throw new Meteor.Error("office365-events:" + error, reason, details);
};