Package.describe({
  name: 'wiseguyeh:office-365-events',
  version: '0.1.2',
  summary: "Provides a fluent interface for managing a user's Office 365 calendar events",
  git: 'https://github.com/djluck/office365-events'
});

Package.onUse(function(api) {
  api.versionsFrom('1.1');
  setupCommonPackageProperties(api);
});

//Package.onTest(function (api) {
//  setupCommonPackageProperties(api);
//  api.use(["sanjo:jasmine"]);
//  api.imply(["sanjo:jasmine"]);
//  api.addFiles("tests/jasmine/server/unit/tests.js", ["client", "server"]);
//});

function setupCommonPackageProperties(api){
  api.use(['wiseguyeh:azure-active-directory@0.3.1', 'underscore@1.0.3', 'check'], 'server');
  api.imply('mrt:moment-timezone@0.2.1');
  api.imply('wiseguyeh:azure-resource-office-365@0.1.1', 'server')
  api.addFiles('office365-events.js');
  api.export("Office365");
}
