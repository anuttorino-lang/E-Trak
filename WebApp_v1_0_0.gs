function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'dashboard';
  var initialAuthToken = e && e.parameter && e.parameter.token ? String(e.parameter.token) : '';
  var fileName = page === 'admin'
    ? 'AdminFullWithProjectCreateAndPhaseTransfer'
    : page === 'login'
      ? 'Login'
      : 'Dashboard';
  var template = HtmlService.createTemplateFromFile(fileName);
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.initialAuthToken = initialAuthToken;

  return template
    .evaluate()
    .setTitle('ระบบติดตามโครงการงานอาคารสถานที่')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppNavigationUrls() {
  var baseUrl = ScriptApp.getService().getUrl();

  return {
    admin: baseUrl + '?page=admin',
    dashboard: baseUrl + '?page=dashboard',
    login: baseUrl + '?page=login',
  };
}
