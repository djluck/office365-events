Office365 = {};
Office365.event = {
    runByUser : function(userOrUserId){
        if (!userOrUSerId) {
            throw new Meteor.Error("office365-events:Missing argument", "userOrUserId", "Please supply a Meteor user or user id");
        }

        if (_.isString(userOrUserId)) {
            userOrUserId = Meteor.users.findOne(userOrUserId);
        }

        return buildFluentObject(user);
    }
}

function buildFluentObject(user){
    var restRequest = {};

    var fluentObject = {
        subject : function(subject){
            restRequest.subject = subject;
            return fluentObject;
        },
        bodyHtml : function(content){
            restRequest.body = {
                ContentType : "HTML",
                Content : content
            };
            return fluentObject;
        },
        bodyPlainText : function(content){
            restRequest.body = {
                ContentType : "Text",
                Content : content
            };
            return fluentObject;
        },
        attendees : function(userIdsOrUsers){
            if (!_.isArray(userIdsOrUsers)){
                throw new Meteor.Error("office365-events:Invalid argument type", "userIdsOrUsers", "Please supply a non-empty array of users or user ids");
            }
            if (_.isEmpty(userIdsOrUsers)){
                throw new Meteor.Error("office365-events:Argument list is empty", "userIdsOrUsers", "Please supply a non-empty array of users or user ids");
            }
            if (_.isString(_.first(userIdsOrUsers))){
                userIdsOrUsers = Meteor.users.find({ _id : { $in : userIdsOrUsers}});
            }
            restRequest.Attendees = _.map(userIdsOrUsers, function(user){
                return {
                    EmailAddress : {
                        Name : user.serviceData.azureAd.displayName,
                        Address : user.serviceData.azureAd.mail || user.serviceData.azureAd.userPrincipleName
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
            restRequest.start = momentTz.format();
            return fluentObject;
        },
        endsAt : function(momentTz){
            checkForMomentTimeZone(momentTz);
            restRequest.end = momentTz.format();
            return fluentObject;
        },
        update : function(eventId){
            var url = "https://outlook.office365.com/api/v1.0/me/events/" + eventId;
            var accessToken = AzureAd.resources.getOrUpdateUserAccessToken(AzureAd.resources.office365.friendlyName, user);
            return AzureAd.http.callAuthenticated("PATCH", url, accessToken, { data : restRequest});
        },
        create : function(){
            var url = "https://outlook.office365.com/api/v1.0/me/events";
            var accessToken = AzureAd.resources.getOrUpdateUserAccessToken(AzureAd.resources.office365.friendlyName, user);
            return AzureAd.http.callAuthenticated("POST", url, accessToken, { data : restRequest});
        },
        delete : function(eventId){
            var url = "https://outlook.office365.com/api/v1.0/me/events/" + eventId;
            var accessToken = AzureAd.resources.getOrUpdateUserAccessToken(AzureAd.resources.office365.friendlyName, user);
            AzureAd.http.callAuthenticated("DELETE", url, accessToken);
        }
    };

    return fluentObject;
}

function checkForMomentTimeZone(obj){
    if (typeof obj !== "Moment" || obj._z === null)
        throw new Meteor.Error("office365-events:Invalid argument type", "momentTz", "Please supply a moment timezone");
}