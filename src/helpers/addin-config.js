/* global Office */

function getConfig() {
  const config = {};
  config.mistralApiKey = Office.context.roamingSettings.get("mistralApiKey");
  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set("mistralApiKey", config.mistralApiKey);
  Office.context.roamingSettings.saveAsync(callback);
}
