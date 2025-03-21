(() => {
  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = (reason) => {
    $(document).ready(() => {
      if (window.location.search) {
        // See if the config values were passed.
        // If so, pre-populate the values.
        const mistralApiKey = getParameterByName("mistralApiKey");
        if (mistralApiKey) {
          $("#mistral-api-key").val(mistralApiKey);
        }
      }

      // When the Done button is selected, send the
      // values back to the caller as a serialized object.
      $("#settings-done").on("click", () => {
        const settings = {};
        settings.mistralApiKey = $("#mistral-api-key").val();

        if (!settings.mistralApiKey) {
          showError("Please enter your Mistral API key");
          return;
        }

        sendMessage(JSON.stringify(settings));
      });
    });
  };

  function showError(error) {
    $("#error-text").text(error);
    $(".error-display").show();
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[[\]]/g, "\\$&");
    const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return "";
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
