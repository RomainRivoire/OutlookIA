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

  // Fonction pour nettoyer la réponse avant insertion (supprimer l'objet suggéré et extraire le texte après "---")
  function cleanResponseForInsertion(response) {
    // Extraire le texte après le délimiteur "---"
    const delimiterIndex = response.indexOf("---");
    let cleanedResponse = delimiterIndex !== -1 ? response.substring(delimiterIndex + 3).trim() : response;

    // Supprimer les lignes contenant des suggestions d'objet
    cleanedResponse = cleanedResponse
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
      // Utiliser directement le format texte avec des sauts de ligne
      // Outlook va automatiquement appliquer le style de texte par défaut du mail
      Office.context.mailbox.item.body.setSelectedDataAsync(
        newResponse,
        { coercionType: Office.CoercionType.Text },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            showError("Impossible d'insérer la réponse: " + (result.error ? result.error.message : "Erreur inconnue"));
            return;
          }

          // Mettre à jour la référence à la dernière insertion
          lastInsertedResponse = newResponse;

          // Appliquer le formatage après insertion
          applyFormattingToLastInserted();
        }
      );
    } catch (e) {
      showError("Erreur lors de l'insertion/remplacement de la réponse: " + e.message);
    }
  }

  // Nouvelle fonction pour appliquer le formatage après insertion
  function applyFormattingToLastInserted() {
    try {
      // Obtenir l'objet Word pour le document
      Office.context.mailbox.item.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Formatted },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log("Impossible d'obtenir la sélection actuelle");
            return;
          }

          // Appliquer le formatage de paragraphe
          Office.context.mailbox.item.body.setSelectedDataAsync(
            result.value,
            {
              coercionType: Office.CoercionType.Text,
              asyncContext: {
                paragraphFormat: {
                  lineSpacing: 1.5,
                  firstLineIndent: 0,
                  alignment: "left",
                  spaceBefore: 12,
                  spaceAfter: 12,
                },
              },
            },
            (result) => {
              console.log("Formatage appliqué");
            }
          );
        }
      );
    } catch (e) {
      console.error("Erreur lors de l'application du formatage:", e);
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
        getEmailContent((emailContent, error) => {
          if (error) {
            $("#ai-loading").hide();
            showError("Erreur lors de la récupération du contenu de l'email: " + error);
            return;
          }

          // Create a prompt with email context and ask for a subject
          const fullPrompt = `
            Je regarde un email avec les détails suivants:
            
            Objet: ${emailContent.subject}
            De: ${emailContent.sender}
            
            Contenu:
            ${emailContent.body}
            
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

      // When insert AI response button is clicked
      $("#insert-ai-response").on("click", () => {
        if (aiResponse) {
          // Nettoyer la réponse avant insertion
          const cleanedResponse = cleanResponseForInsertion(aiResponse);

          // Remplacer la dernière insertion ou insérer une nouvelle réponse
          replaceOrInsertResponse(cleanedResponse);
        }
      });

      // When insert subject button is clicked
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
