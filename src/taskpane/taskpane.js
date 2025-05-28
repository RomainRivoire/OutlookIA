/* eslint-disable office-addins/no-office-initialize */
/* eslint-disable no-undef */
(() => {
  /* global Office */
  /* global $ */

  let config;
  let aiResponse = "";
  let generatedSubject = "";
  let lastInsertedResponse = ""; // Pour suivre la dernière réponse insérée
  const qaHistory = [];
  // Function to log messages
  function logMessage(message) {
    console.log(message); // Log to console
    // Optionally, log to a specific area in the task pane
    $("#log-area").append(`<div>${new Date().toISOString()}: ${message}</div>`);
  }

  // Function to get email content with logging
  function getEmailContent(callback) {
    try {
      logMessage("Getting email content...");
      const item = Office.context.mailbox.item;
      const conversationId = item.conversationId;

      logMessage("REST URL:", Office.context.mailbox.restUrl);

      logMessage("Exchange REST disponible ? " + (Office.context.mailbox.restUrl ? "Oui" : "Non"));

      logMessage(`REST URL réelle: ${Office.context.mailbox.restUrl}`);
      logMessage(`EWS URL: ${Office.context.mailbox.ewsUrl}`);

      Office.context.mailbox.getCallbackTokenAsync(
        {
          forceConsent: false,
          isRest: true,
          scopes: ["mail.read"],
          diagnostics: {
            traceFlags: 63, // Niveau de trace maximal
            customData: JSON.stringify({
              clientVersion: Office.context.diagnostics.version,
              mailboxType: Office.context.mailbox.diagnostics.mailboxType,
            }),
          },
        },
        function (result) {
          logMessage("Call graph :" + result.status);
          if (result.status === Office.AsyncResultStatus.Failed) {
            logMessage(`Error code: ${result.error.code}, Name: ${result.error.name}`);
            logMessage("Diagnostics complets:" + result.error.httpStatus);
          }

          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const accessToken = result.value;
            logMessage("Access token retrieved successfully.");

            $.ajax({
              url: `https://graph.microsoft.com/v2.0/me/messages?$filter=conversationId eq '${conversationId}'`,
              type: "GET",
              headers: {
                Authorization: `Bearer ${accessToken}`,
              },
            })
              .done(function (response) {
                if (response.value && response.value.length > 0) {
                  const emails = response.value;
                  const emailContents = emails.map((email) => ({
                    subject: email.subject || "",
                    sender: email.from ? email.from.emailAddress.address : "Inconnu",
                    body: email.body.content,
                  }));

                  logMessage("Email content retrieved successfully.");
                  callback(emailContents);
                } else {
                  logMessage("No conversation history found.");
                  callback(null, "Aucun historique de conversation trouvé");
                }
              })
              .fail(function (error) {
                logMessage("Error retrieving conversation history: " + error.message);
                callback(null, "Erreur lors de la récupération de l'historique de conversation: " + error.message);
              });
          } else {
            logMessage("Error retrieving access token: " + result.error.message);
            callback(null, "Erreur lors de la récupération du token d'accès: " + result.error.message);
          }
        }
      );
    } catch (e) {
      logMessage("Error accessing email: " + e.message);
      callback(null, "Erreur lors de l'accès à l'email: " + e.message);
    }
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

  // Fonction pour extraire l'objet généré de la réponse
  function extractSubject(response) {
    // Recherche d'un objet suggéré dans la réponse
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

    // Si aucun objet n'est trouvé, générer un objet par défaut
    return "Réponse générée par IA";
  }

  // Fonction pour nettoyer la réponse avant insertion (supprimer l'objet suggéré)
  function cleanResponseForInsertion(response) {
    // Supprimer les lignes contenant des suggestions d'objet
    let cleanedResponse = response
      .replace(/Objet suggéré\s*:.*(\n|$)/gi, "")
      .replace(/Sujet suggéré\s*:.*(\n|$)/gi, "")
      .replace(/Objet\s*:.*(\n|$)/gi, "")
      .replace(/Sujet\s*:.*(\n|$)/gi, "")
      .trim();

    // Supprimer les formules d'introduction courantes
    cleanedResponse = cleanedResponse
      .replace(
        /^(Bien sûr|Voici|Certainement|Avec plaisir|D'accord|Bonjour|Salut|Je serais ravi)[,.]?\s*(voici|je vous propose|vous trouverez ci-dessous)?\s*(une réponse possible|une réponse|un exemple de réponse|ma réponse)[^:]*:/i,
        ""
      )
      .replace(/^Voici ma réponse[^:]*:/i, "")
      .replace(/^En réponse à votre demande[^:]*:/i, "")
      .replace(/^Voici ce que vous pourriez répondre[^:]*:/i, "")
      .trim();

    // Supprimer les marqueurs de formatage
    cleanedResponse = cleanedResponse
      .replace(/^-{3,}$/gm, "") // Supprime les lignes contenant uniquement des tirets (---)
      .replace(/^\*{3,}$/gm, "") // Supprime les lignes contenant uniquement des astérisques (***)
      .trim();

    // Supprimer les placeholders courants
    cleanedResponse = cleanedResponse
      .replace(/\[Votre Nom\]/gi, "")
      .replace(/\[Votre Signature\]/gi, "")
      .replace(/\[Nom\]/gi, "")
      .replace(/\[Prénom\]/gi, "")
      .replace(/\[Signature\]/gi, "")
      .trim();

    // Améliorer le formatage des sauts de ligne
    cleanedResponse = cleanedResponse
      .replace(/\n{3,}/g, "\n\n") // Remplace 3+ sauts de ligne par 2
      .replace(/^\s+|\s+$/gm, ""); // Supprime les espaces en début et fin de ligne

    // Assurer que les paragraphes sont bien séparés pour Outlook
    cleanedResponse = cleanedResponse
      .replace(/\n/g, "\r\n") // Utiliser le format Windows pour les sauts de ligne
      .replace(/([.!?])\s*\r\n/g, "$1\r\n\r\n"); // Ajouter un saut de ligne supplémentaire après les fins de phrase suivies d'un saut de ligne

    return cleanedResponse;
  }

  // Fonction pour définir l'objet du mail
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

  // Fonction pour remplacer la dernière insertion ou insérer une nouvelle réponse
  function replaceOrInsertResponse(newResponse) {
    try {
      // Récupérer le contenu actuel du mail
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showError(
            "Impossible de lire le contenu du mail: " + (result.error ? result.error.message : "Erreur inconnue")
          );
          return;
        }

        const currentBody = result.value;

        // Si nous avons une dernière insertion et qu'elle est présente dans le corps du mail
        if (lastInsertedResponse && currentBody.includes(lastInsertedResponse)) {
          // Remplacer la dernière insertion par la nouvelle
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

              // Mettre à jour la référence à la dernière insertion
              lastInsertedResponse = newResponse;
            }
          );
        } else {
          // Si pas de dernière insertion ou si elle n'est plus présente, insérer à la position actuelle
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

              // Mettre à jour la référence à la dernière insertion
              lastInsertedResponse = newResponse;
            }
          );
        }
      });
    } catch (e) {
      showError("Erreur lors de l'insertion/remplacement de la réponse: " + e.message);
    }
  }

  // Nouvelle fonction pour ajouter une paire question/réponse à l'historique
  function addToQAHistory(question, response) {
    // Ajouter à la fin du tableau pour garder la dernière question en bas (comme WhatsApp)
    qaHistory.push({ question, response });

    // Mettre à jour l'affichage
    updateQAHistoryDisplay();
  }

  // Nouvelle fonction pour mettre à jour l'affichage de l'historique
  function updateQAHistoryDisplay() {
    const $qaHistory = $("#qa-history");
    $qaHistory.empty();

    // Parcourir l'historique et créer les éléments HTML
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

    // Faire défiler vers le bas pour voir la dernière question/réponse
    $(".conversation-container").scrollTop($(".conversation-container")[0].scrollHeight);
  }

  Office.initialize = (reason) => {
    $(document).ready(() => {
      // Load config
      config = getConfig();

      // Vérifier si la clé API est déjà configurée
      if (!config || !config.mistralApiKey) {
        // Si la clé n'est pas configurée, afficher un message d'erreur
        showError("Clé API Mistral non configurée. Veuillez contacter votre administrateur.");
        return;
      }

      // When AI submit button is clicked
      $("#ai-submit").on("click", () => {
        const prompt = $("#ai-prompt").val().trim();
        if (!prompt) {
          showError("Veuillez entrer une question sur cet email");
          return;
        }

        // Hide any previous errors
        $("#error-display").hide();

        // Hide action buttons
        $("#action-buttons").hide();

        // Show loading spinner
        $("#ai-loading").show();

        // Get email content
        getEmailContent((emailContents, error) => {
          if (error) {
            $("#ai-loading").hide();
            showError("Erreur lors de la récupération du contenu de l'email: " + error);
            return;
          }

          // Create a prompt with email context and ask for a subject
          const fullPrompt = `
            Je regarde une conversation avec les détails suivants:

            ${emailContents
              .map(
                (emailContent, index) => `
                Email ${index + 1}:
                Objet: ${emailContent.subject}
                De: ${emailContent.sender}

                Contenu:
                ${emailContent.body}
              `
              )
              .join("\n\n")}

            ${prompt}

            En plus de répondre à ma question, pourriez-vous également suggérer un objet approprié pour ma réponse? Présentez-le sous la forme "Objet suggéré: [votre suggestion d'objet]" à la fin de votre réponse.

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

            // Sauvegarder la réponse complète
            aiResponse = response;

            // Extraire et sauvegarder l'objet suggéré
            generatedSubject = extractSubject(response);
            console.log("Objet généré:", generatedSubject);

            // Ajouter la question et la réponse à l'historique
            addToQAHistory(prompt, response);

            // Vider la zone de prompt
            $("#ai-prompt").val("");

            // Afficher les boutons d'action
            $("#action-buttons").show();
          });
        });
      });

      $("#insert-ai-response").on("click", () => {
        if (aiResponse) {
          // Nettoyer la réponse avant insertion
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
