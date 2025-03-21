(() => {
  /* global Office */
  /* global $ */

  let config;
  let aiResponse = "";

  // Add the missing functions inside the scope
  function getEmailContent(callback) {
    try {
      // Get the current item (email)
      const item = Office.context.mailbox.item;

      // Get the email body
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = result.value;
          const subject = item.subject || "";

          // Get the sender
          const sender = item.sender;
          const senderEmail = sender ? sender.emailAddress : "Inconnu";

          // Combine all information
          const emailContent = {
            subject: subject,
            sender: senderEmail,
            body: emailBody,
          };

          callback(emailContent);
        } else {
          callback(
            null,
            result.error ? result.error.message : "Erreur inconnue lors de la récupération du contenu de l'email"
          );
        }
      });
    } catch (e) {
      callback(null, "Erreur lors de l'accès à l'email: " + e.message);
    }
  }

  function getConfig() {
    // Implementation for getConfig
    // This depends on how you're storing configuration
    // Example implementation:
    try {
      const savedConfig = Office.context.roamingSettings.get('mistralConfig');
      return savedConfig ? JSON.parse(savedConfig) : {};
    } catch (e) {
      console.error("Erreur lors de la récupération de la configuration:", e);
      return {};
    }
  }

  function setConfig(configObj, callback) {
    try {
      Office.context.roamingSettings.set('mistralConfig', JSON.stringify(configObj));
      Office.context.roamingSettings.saveAsync((result) => {
        callback(result);
      });
    } catch (e) {
      callback({
        status: Office.AsyncResultStatus.Failed,
        error: { message: e.message }
      });
    }
  }

  function callMistralAPI(apiKey, prompt, callback) {
    const requestUrl = "https://api.mistral.ai/v1/chat/completions";
    
    const requestData = {
      model: "mistral-small-latest",
      messages: [
        {
          role: "user",
          content: prompt,
        },
      ],
      temperature: 0.7,
      max_tokens: 1024,
    };
    
    $.ajax({
      url: requestUrl,
      type: "POST",
      dataType: "json",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + apiKey,
      },
      data: JSON.stringify(requestData),
    })
      .done((response) => {
        if (response && response.choices && response.choices.length > 0) {
          callback(response.choices[0].message.content);
        } else {
          callback(null, "Réponse invalide de l'API Mistral");
        }
      })
      .fail((error) => {
        console.error("Erreur API Mistral:", error);
        let errorMessage = "Erreur lors de l'appel à l'API Mistral";
        
        if (error.status === 401) {
          errorMessage = "Clé API invalide. Veuillez vérifier votre clé API Mistral.";
        } else if (error.responseJSON && error.responseJSON.error) {
          errorMessage += ": " + error.responseJSON.error.message;
        } else if (error.statusText) {
          errorMessage += ": " + error.statusText;
        }
        
        callback(null, errorMessage);
      });
  }

  function showError(error) {
    $("#error-message").text(error);
    $("#error-display").show();
  }

  Office.initialize = (reason) => {
    $(document).ready(() => {
      // Load config
      config = getConfig();

      // If API key is already saved, show it in the input field and show AI container
      if (config && config.mistralApiKey) {
        $("#mistral-api-key").val(config.mistralApiKey);
        $("#ai-container").show();
      }

      // Save API key button click handler
      $("#save-api-key").on("click", () => {
        const apiKey = $("#mistral-api-key").val().trim();

        if (!apiKey) {
          showError("Veuillez entrer votre clé API Mistral");
          return;
        }

        // Save API key to config
        config = config || {};
        config.mistralApiKey = apiKey;

        setConfig(config, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            showError("Erreur lors de l'enregistrement de la clé API: " + result.error.message);
            return;
          }

          // Show AI container
          $("#ai-container").show();

          // Hide error if any
          $("#error-display").hide();
        });
      });

      // When AI submit button is clicked
      $("#ai-submit").on("click", () => {
        const prompt = $("#ai-prompt").val().trim();
        if (!prompt) {
          showError("Veuillez entrer une question sur cet email");
          return;
        }

        // Check if API key is available
        if (!config || !config.mistralApiKey) {
          showError("Veuillez d'abord enregistrer votre clé API Mistral");
          return;
        }

        // Hide any previous errors
        $("#error-display").hide();

        // Show loading spinner
        $("#ai-loading").show();
        $("#ai-response").hide();
        $("#insert-ai-response").hide();

        // Get email content
        getEmailContent((emailContent, error) => {
          if (error) {
            $("#ai-loading").hide();
            showError("Erreur lors de la récupération du contenu de l'email: " + error);
            return;
          }

          // Create a prompt with email context
          const fullPrompt = `
            Je regarde un email avec les détails suivants:
            
            Objet: ${emailContent.subject}
            De: ${emailContent.sender}
            
            Contenu:
            ${emailContent.body}
            
            ${prompt}
            
            Veuillez répondre en français.
          `;

          console.log("Appel de l'API Mistral...");

          // Call Mistral API
          callMistralAPI(config.mistralApiKey, fullPrompt, (response, error) => {
            $("#ai-loading").hide();

            if (error) {
              showError(error);
              return;
            }

            console.log("Réponse reçue de l'API Mistral");

            // Display the response
            $("#ai-response").show().find(".response-content").html(response.replace(/\n/g, "<br>"));
            aiResponse = response;
            $("#insert-ai-response").show();
          });
        });
      });

      // When insert AI response button is clicked
      $("#insert-ai-response").on("click", () => {
        if (aiResponse) {
          Office.context.mailbox.item.body.setSelectedDataAsync(
            aiResponse,
            { coercionType: Office.CoercionType.Text },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Failed) {
                showError("Impossible d'insérer la réponse de l'IA: " + result.error.message);
              }
            }
          );
        }
      });
    });
  };
})();