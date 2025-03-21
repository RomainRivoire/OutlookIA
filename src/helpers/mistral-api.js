import $ from "jquery";
import * as Office from "office-js";

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
