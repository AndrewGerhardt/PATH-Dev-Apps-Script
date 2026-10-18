function isRunActive_() {
  return PropertiesService.getScriptProperties().getProperty("RUN_ACTIVE") === "true";
}

function setRunActive_(value) {
  PropertiesService.getScriptProperties().setProperty("RUN_ACTIVE", value ? "true" : "false");
}