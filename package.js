Package.describe({
  name: 'wiseguyeh:office-365-events',
  version: '0.1.3',
  summary: "Provides a fluent interface for managing a user's Office 365 calendar events",
  git: 'https://github.com/djluck/office365-events'
});

Package.onUse(function(api) {
  api.versionsFrom('1.1');
  setupCommon(api);
});

Package.onTest(function (api) {
  setupCommon(api);
  api.use(["tinytest", "accounts-password", "practicalmeteor:sinon@1.10.3_2"]);
  api.imply(["tinytest", "accounts-password", "practicalmeteor:sinon@1.10.3_2", "mrt:moment-timezone@0.2.1"]);
  api.addFiles([
        "tests/lib.js",
        "tests/tests.js"
      ],
      "server"
  );
});

function setupCommon(api){
  api.use(['wiseguyeh:azure-active-directory@0.3.1', 'underscore@1.0.3'], 'server');
  api.imply('wiseguyeh:azure-resource-office-365@0.1.1', 'server')
  api.addFiles('office365-events.js', 'server');
  api.export("Office365", "server");
}
