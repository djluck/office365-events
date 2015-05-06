
var originalAdd = Tinytest.add;
Tinytest.add = function(name, fn){
    originalAdd(name, function(test){
        test.throwsException = throwsException.bind(this, test);
        fn(test);
    })
}

Tinytest.addWithSinon = function(name, fnToRun){
    return Tinytest.add(name, function(test){
        var testContext = this;
        sinon.test(function(){
            var sinonSanbox = this;
            test.assertOffice365FluentObjectContains = assertOffice365FluentObjectContains.bind(testContext, test, sinonSanbox);
            sinonSanbox.arrangeAccessTokenForUser = arrangeAccessTokenForUser.bind(testContext, sinonSanbox);
            sinonSanbox.arrangeExpectedHttpCall = arrangeExpectedHttpCall.bind(testContext, sinonSanbox);
            fnToRun.call(testContext, test, this);
        })();
    })
}

arrangeUser = function(user){
    var id = Meteor.users.insert(user || {});

    return Meteor.users.findOne(id);
}

arrangeUserWithAzureCredentials = function(email, displayName){
    return arrangeUser({
        services : {
            azureAd : {
                displayName : displayName || "dummy display name",
                mail : email || "test@dummymail.com"
            }
        }
    })
}

function throwsException(test, fn, expectedError){
    var thrown = false;
    try{
        fn();
    }
    catch(ex){
        thrown = true;
        test.isTrue(ex && "error" in ex, "Exception did not have error message, was: " + ex);
        test.equal(ex.error, "office365-events:" + expectedError, "Error message did not match expected, was: " + ex.error);
    }
    test.isTrue(thrown, "No exception was thrown");
}

function arrangeAccessTokenForUser(sinonSandbox, user, accessToken){
    sinonSandbox.stub(AzureAd.resources, "getOrUpdateUserAccessToken")
        .withArgs(AzureAd.resources.office365.friendlyName, user)
        .returns(accessToken);
}

function arrangeExpectedHttpCall(sinonSandbox, madeByUser, url, verb, expectedData, returns){
    var accessToken = "anAccessToken";
    sinonSandbox.arrangeAccessTokenForUser(madeByUser, accessToken);

    sinonSandbox.mock(AzureAd.http)
        .expects("callAuthenticated")
        .withArgs(verb, url, accessToken, !expectedData ?  undefined : { data : expectedData })
        .returns(returns || { Id : 1});
}

function assertOffice365FluentObjectContains(test, sinonSandbox, fluentObject, eventOwner, expectedApiBody){
    sinonSandbox.arrangeExpectedHttpCall(eventOwner, "https://outlook.office365.com/api/v1.0/me/events", "POST", expectedApiBody);
    fluentObject.create();
}