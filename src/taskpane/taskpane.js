/* eslint-disable office-addins/no-office-initialize */
/* eslint-disable no-undef */
const GraphHelper = require("../helpers/graphHelper.js").default;

(() => {
  /* global Office */
  /* global $ */

  let config;
  let aiResponse = "";
  let generatedSubject = "";
  let lastInsertedResponse = "";
  const qaHistory = [];
  const graphHelper = new GraphHelper();

  // Function to log messages
  function logMessage(message) {
    console.log(message);
    $("#log-area").append(`<div>${new Date().toISOString()}: ${message}</div>`);
  }

  async function getEmailContent(callback) {
    try {
      logMessage("Getting email content...");
      const item = Office.context.mailbox.item;
      const conversationId = item.conversationId;

      logMessage(`Conversation ID: ${conversationId}`);

      // Use await to get the access token
      logMessage("test");

      const access_Token = await graphHelper.getAccessToken();
      logMessage("token : " + access_Token);

      // Essayer d'abord avec l'API REST moderne
      tryRestAPIAccess(conversationId, callback);
    } catch (e) {
      logMessage("Error accessing email: " + e.message);
      callback(null, "Erreur lors de l'accès à l'email: " + e.message);
    }
  }

  // Fonction pour essayer l'accès via REST API
  async function tryRestAPIAccess(conversationId, callback) {
    try {
      const accessToken = await graphHelper.getAccessToken();
      logMessage("REST token retrieved successfully");

      // Appeler Microsoft Graph API
      callGraphAPI(conversationId, accessToken, callback);
    } catch (error) {
      logMessage(`REST token error: ${error.message}`);

      // Fallback: utiliser seulement l'email courant
      getCurrentEmailOnly(callback);
    }
  }

  // Fonction pour appeler Microsoft Graph API
  function callGraphAPI(conversationId, accessToken, callback) {
    const graphUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'&$expand=attachments&$select=id,subject,from,body,receivedDateTime,attachments`;

    logMessage(`Calling Graph API: ${graphUrl}`);

    $.ajax({
      url: graphUrl,
      type: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      timeout: 30000,
    })
      .done(function (response) {
        logMessage("Graph API call successful");

        if (response.value && response.value.length > 0) {
          const emails = response.value;
          const emailContents = emails.map((email) => ({
            id: email.id,
            subject: email.subject || "Sans objet",
            sender: email.from ? email.from.emailAddress.address : "Expéditeur inconnu",
            senderName: email.from ? email.from.emailAddress.name : "Nom inconnu",
            body: email.body.content || "",
            bodyType: email.body.contentType || "text",
            receivedDateTime: email.receivedDateTime,
            attachments: email.attachments || [],
          }));

          logMessage(`Retrieved ${emailContents.length} emails from conversation`);

          // Log des pièces jointes trouvées
          emailContents.forEach((email, index) => {
            if (email.attachments.length > 0) {
              logMessage(`Email ${index + 1} has ${email.attachments.length} attachments`);
              email.attachments.forEach((att) => {
                logMessage(`- Attachment: ${att.name} (${att.contentType})`);
              });
            }
          });

          callback(emailContents);
        } else {
          logMessage("No emails found in conversation");
          callback(null, "Aucun email trouvé dans cette conversation");
        }
      })
      .fail(function (xhr, status, error) {
        logMessage(`Graph API error: ${status} - ${error}`);
        logMessage(`Response: ${xhr.responseText}`);

        // Fallback: utiliser seulement l'email courant
        getCurrentEmailOnly(callback);
      });
  }

  // Fonction de fallback : récupérer seulement l'email courant
  function getCurrentEmailOnly(callback) {
    logMessage("Using current email only as fallback...");

    try {
      const item = Office.context.mailbox.item;

      // Récupérer le corps de l'email courant
      item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const emailContent = [
            {
              id: item.itemId || "current",
              subject: item.subject || "Sans objet",
              sender: item.from ? item.from.emailAddress : "Expéditeur inconnu",
              senderName: item.from ? item.from.displayName : "Nom inconnu",
              body: result.value,
              bodyType: "html",
              receivedDateTime: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString(),
              attachments: item.attachments || [],
            },
          ];

          logMessage("Current email content retrieved successfully");
          callback(emailContent);
        } else {
          callback(null, "Impossible de récupérer le contenu de l'email courant");
        }
      });
    } catch (e) {
      callback(null, `Erreur lors de la récupération de l'email courant: ${e.message}`);
    }
  }

  // Fonction pour traiter les pièces jointes
  function processAttachments(attachments) {
    if (!attachments || attachments.length === 0) {
      return "Aucune pièce jointe";
    }

    return attachments
      .map((att) => {
        return `- ${att.name} (${att.contentType || "type inconnu"}, ${att.size || "taille inconnue"} octets)`;
      })
      .join("\n");
  }

  function getConfig() {
    try {
      const savedConfig = Office.context.roamingSettings.get("mistralConfig");
      return savedConfig ? JSON.parse(savedConfig) : {};
    } catch (e) {
      console.error("Erreur lors de la récupération de la configuration:", e);
      return {};
    }
  }

  function setConfig(configObj, callback) {
    try {
      Office.context.roamingSettings.set("mistralConfig", JSON.stringify(configObj));
      Office.context.roamingSettings.saveAsync((result) => {
        callback(result);
      });
    } catch (e) {
      callback({
        status: Office.AsyncResultStatus.Failed,
        error: { message: e.message },
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

  function extractSubject(response) {
    const subjectPatterns = [
      /Objet suggéré\s*:\s*"([^"]+)"/i,
      /Objet suggéré\s*:\s*(.+?)(?:\n|$)/i,
      /Objet\s*:\s*"([^"]+)"/i,
      /Objet\s*:\s*(.+?)(?:\n|$)/i,
      /Sujet suggéré\s*:\s*"([^"]+)"/i,
      /Sujet suggéré\s*:\s*(.+?)(?:\n|$)/i,
      /Sujet\s*:\s*"([^"]+)"/i,
      /Sujet\s*:\s*(.+?)(?:\n|$)/i,
    ];

    for (const pattern of subjectPatterns) {
      const match = response.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }

    return "Réponse générée par IA";
  }

  function cleanResponseForInsertion(response) {
    let cleanedResponse = response
      .replace(/Objet suggéré\s*:.*(\n|$)/gi, "")
      .replace(/Sujet suggéré\s*:.*(\n|$)/gi, "")
      .replace(/Objet\s*:.*(\n|$)/gi, "")
      .replace(/Sujet\s*:.*(\n|$)/gi, "")
      .trim();

    cleanedResponse = cleanedResponse
      .replace(
        /^(Bien sûr|Voici|Certainement|Avec plaisir|D'accord|Bonjour|Salut|Je serais ravi)[,.]?\s*(voici|je vous propose|vous trouverez ci-dessous)?\s*(une réponse possible|une réponse|un exemple de réponse|ma réponse)[^:]*:/i,
        ""
      )
      .replace(/^Voici ma réponse[^:]*:/i, "")
      .replace(/^En réponse à votre demande[^:]*:/i, "")
      .replace(/^Voici ce que vous pourriez répondre[^:]*:/i, "")
      .trim();

    cleanedResponse = cleanedResponse
      .replace(/^-{3,}$/gm, "")
      .replace(/^\*{3,}$/gm, "")
      .trim();

    cleanedResponse = cleanedResponse
      .replace(/\[Votre Nom\]/gi, "")
      .replace(/\[Votre Signature\]/gi, "")
      .replace(/\[Nom\]/gi, "")
      .replace(/\[Prénom\]/gi, "")
      .replace(/\[Signature\]/gi, "")
      .trim();

    cleanedResponse = cleanedResponse.replace(/\n{3,}/g, "\n\n").replace(/^\s+|\s+$/gm, "");

    cleanedResponse = cleanedResponse.replace(/\n/g, "\r\n").replace(/([.!?])\s*\r\n/g, "$1\r\n\r\n");

    return cleanedResponse;
  }

  function setEmailSubject(subject) {
    try {
      Office.context.mailbox.item.subject.setAsync(subject, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showError("Impossible de définir l'objet: " + (result.error ? result.error.message : "Erreur inconnue"));
        }
      });
    } catch (e) {
      showError("Erreur lors de la définition de l'objet: " + e.message);
    }
  }

  function replaceOrInsertResponse(newResponse) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showError(
            "Impossible de lire le contenu du mail: " + (result.error ? result.error.message : "Erreur inconnue")
          );
          return;
        }

        const currentBody = result.value;

        if (lastInsertedResponse && currentBody.includes(lastInsertedResponse)) {
          const updatedBody = currentBody.replace(lastInsertedResponse, newResponse);

          Office.context.mailbox.item.body.setAsync(
            updatedBody,
            { coercionType: Office.CoercionType.Text },
            (result) => {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                showError(
                  "Impossible de mettre à jour la réponse: " + (result.error ? result.error.message : "Erreur inconnue")
                );
                return;
              }

              lastInsertedResponse = newResponse;
            }
          );
        } else {
          Office.context.mailbox.item.body.setSelectedDataAsync(
            newResponse,
            { coercionType: Office.CoercionType.Text },
            (result) => {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                showError(
                  "Impossible d'insérer la réponse: " + (result.error ? result.error.message : "Erreur inconnue")
                );
                return;
              }

              lastInsertedResponse = newResponse;
            }
          );
        }
      });
    } catch (e) {
      showError("Erreur lors de l'insertion/remplacement de la réponse: " + e.message);
    }
  }

  function addToQAHistory(question, response) {
    qaHistory.push({ question, response });
    updateQAHistoryDisplay();
  }

  function updateQAHistoryDisplay() {
    const $qaHistory = $("#qa-history");
    $qaHistory.empty();

    qaHistory.forEach((item) => {
      const $qaItem = $(`
        <div class="qa-item">
          <div class="question-container">
            <div class="question-header">Votre question:</div>
            <p class="question-content">${item.question}</p>
          </div>
          <div class="response-container">
            <div class="response-header">Réponse de l'IA:</div>
            <div class="response-content">${item.response.replace(/\n/g, "<br>")}</div>
          </div>
        </div>
      `);

      $qaHistory.append($qaItem);
    });

    $(".conversation-container").scrollTop($(".conversation-container")[0].scrollHeight);
  }

  Office.initialize = (reason) => {
    $(document).ready(() => {
      config = getConfig();

      if (!config || !config.mistralApiKey) {
        showError("Clé API Mistral non configurée. Veuillez contacter votre administrateur.");
        return;
      }

      $("#ai-submit").on("click", () => {
        const prompt = $("#ai-prompt").val().trim();
        if (!prompt) {
          showError("Veuillez entrer une question sur cet email");
          return;
        }

        $("#error-display").hide();
        $("#action-buttons").hide();
        $("#ai-loading").show();

        getEmailContent((emailContents, error) => {
          if (error) {
            $("#ai-loading").hide();
            showError("Erreur lors de la récupération du contenu de l'email: " + error);
            return;
          }

          // Créer un prompt enrichi avec les informations sur les pièces jointes
          const fullPrompt = `
            Je regarde une conversation avec les détails suivants:

            ${emailContents
              .map(
                (emailContent, index) => `
                Email ${index + 1}:
                Objet: ${emailContent.subject}
                De: ${emailContent.senderName} (${emailContent.sender})
                Date: ${emailContent.receivedDateTime}

                Contenu:
                ${emailContent.body}

                Pièces jointes:
                ${processAttachments(emailContent.attachments)}
              `
              )
              .join("\n\n")}

            ${prompt}

            En plus de répondre à ma question, pourriez-vous également suggérer un objet approprié pour ma réponse? Présentez-le sous la forme "Objet suggéré: [votre suggestion d'objet]" à la fin de votre réponse.

            Veuillez répondre en français.
          `;

          console.log("Appel de l'API Mistral...");

          callMistralAPI(config.mistralApiKey, fullPrompt, (response, error) => {
            $("#ai-loading").hide();

            if (error) {
              showError(error);
              return;
            }

            console.log("Réponse reçue de l'API Mistral");

            aiResponse = response;
            generatedSubject = extractSubject(response);
            console.log("Objet généré:", generatedSubject);

            addToQAHistory(prompt, response);
            $("#ai-prompt").val("");
            $("#action-buttons").show();
          });
        });
      });

      $("#insert-ai-response").on("click", () => {
        if (aiResponse) {
          const cleanedResponse = cleanResponseForInsertion(aiResponse);
          replaceOrInsertResponse(cleanedResponse);
        }
      });

      $("#insert-subject").on("click", () => {
        if (generatedSubject) {
          setEmailSubject(generatedSubject);
        } else {
          showError("Aucun objet n'a été généré");
        }
      });
    });
  };
})();
