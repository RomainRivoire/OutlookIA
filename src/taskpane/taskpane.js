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

      // Essayer d'abord avec l'API REST
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
      callGraphAPI(conversationId, accessToken, callback);
    } catch (error) {
      logMessage(`REST token error: ${error.message}`);
      // Fallback: utiliser seulement l'email courant
      getCurrentEmailOnly(callback);
    }
  }

  // Fonction pour vérifier si une pièce jointe est PDF ou DOCX
  function isRelevantAttachment(attachment) {
    const relevantTypes = [
      "application/pdf",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/msword",
    ];

    const fileExtensions = [".pdf", ".docx", ".doc"];

    // Vérifier par type MIME
    if (attachment.contentType && relevantTypes.includes(attachment.contentType.toLowerCase())) {
      return true;
    }

    // Vérifier par extension de fichier
    if (attachment.name) {
      const fileName = attachment.name.toLowerCase();
      return fileExtensions.some((ext) => fileName.endsWith(ext));
    }

    return false;
  }

  // Fonction pour lire le contenu d'une pièce jointe
  async function readAttachmentContent(attachment, accessToken, userEmail) {
    if (!isRelevantAttachment(attachment)) {
      return null;
    }

    try {
      logMessage(`Reading attachment: ${attachment.name}`);

      // URL pour récupérer le contenu de la pièce jointe
      const attachmentUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userEmail)}/messages/${attachment.messageId}/attachments/${attachment.id}/$value`;

      const response = await fetch(attachmentUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (!response.ok) {
        logMessage(`Failed to read attachment ${attachment.name}: ${response.status}`);
        return {
          name: attachment.name,
          type: attachment.contentType || "unknown",
          content: "[Contenu non accessible]",
          error: `Erreur ${response.status}`,
          size: attachment.size || 0,
        };
      }

      const arrayBuffer = await response.arrayBuffer();

      // Pour les PDF, on ne peut pas extraire le texte directement côté client
      // On indique simplement la présence du fichier
      if (attachment.contentType === "application/pdf" || attachment.name.toLowerCase().endsWith(".pdf")) {
        return {
          name: attachment.name,
          type: "PDF",
          content: "[Document PDF présent - contenu non extrait]",
          size: arrayBuffer.byteLength,
        };
      }

      // Pour les documents Word (.docx), on peut essayer d'extraire le texte
      if (
        attachment.contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
        attachment.name.toLowerCase().endsWith(".docx")
      ) {
        // Ici, vous pourriez utiliser une bibliothèque comme mammoth.js pour extraire le texte
        // Pour l'instant, on indique simplement la présence du document
        return {
          name: attachment.name,
          type: "DOCX",
          content: "[Document Word présent - contenu non extrait]",
          size: arrayBuffer.byteLength,
        };
      }

      // Pour les anciens documents Word (.doc)
      if (attachment.contentType === "application/msword" || attachment.name.toLowerCase().endsWith(".doc")) {
        return {
          name: attachment.name,
          type: "DOC",
          content: "[Document Word (ancien format) présent - contenu non extrait]",
          size: arrayBuffer.byteLength,
        };
      }

      return {
        name: attachment.name,
        type: attachment.contentType || "unknown",
        content: "[Document présent]",
        size: arrayBuffer.byteLength,
      };
    } catch (error) {
      logMessage(`Error reading attachment ${attachment.name}: ${error.message}`);
      return {
        name: attachment.name,
        type: attachment.contentType || "unknown",
        content: "[Erreur lors de la lecture]",
        error: error.message,
        size: attachment.size || 0,
      };
    }
  }

  // Fonction pour appeler Microsoft Graph API
  async function callGraphAPI(conversationId, accessToken, callback) {
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    if (!userEmail) {
      logMessage("Unable to get user email from Office context");
      getCurrentEmailOnly(callback);
      return;
    }

    logMessage(`Using user email: ${userEmail}`);

    try {
      // Étape 1: Récupérer la liste des emails de la conversation (sans pièces jointes)
      const messagesUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
        userEmail
      )}/messages?$filter=conversationId eq '${encodeURIComponent(conversationId)}'&$select=id,subject,from,body`;

      logMessage(`Step 1: Getting messages list: ${messagesUrl}`);

      const messagesResponse = await fetch(messagesUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!messagesResponse.ok) {
        const errorText = await messagesResponse.text();
        logMessage(`Messages API error: ${messagesResponse.status} - ${messagesResponse.statusText}`);
        logMessage(`Response: ${errorText}`);
        getCurrentEmailOnly(callback);
        return;
      }

      const messagesData = await messagesResponse.json();
      logMessage("Messages API call successful");

      if (!messagesData.value || messagesData.value.length === 0) {
        logMessage("No emails found in conversation");
        callback(null, "Aucun email trouvé dans cette conversation");
        return;
      }

      const emails = messagesData.value;
      logMessage(`Found ${emails.length} emails in conversation`);

      const emailContents = [];

      // Étape 2: Pour chaque email, récupérer ses pièces jointes séparément
      for (const email of emails) {
        logMessage(`Processing email: ${email.subject}`);

        // Récupérer les pièces jointes pour cet email spécifique
        const attachmentsUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
          userEmail
        )}/messages/${email.id}/attachments?$select=id,name,contentType,size`;

        let attachments = [];
        try {
          const attachmentsResponse = await fetch(attachmentsUrl, {
            method: "GET",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          });

          if (attachmentsResponse.ok) {
            const attachmentsData = await attachmentsResponse.json();
            attachments = attachmentsData.value || [];
            logMessage(`Found ${attachments.length} attachments for email ${email.subject}`);
          } else {
            logMessage(`Failed to get attachments for email ${email.subject}: ${attachmentsResponse.status}`);
          }
        } catch (attachError) {
          logMessage(`Error getting attachments for email ${email.subject}: ${attachError.message}`);
        }

        // Filtrer les pièces jointes pertinentes
        const relevantAttachments = attachments.filter((att) => isRelevantAttachment(att));
        logMessage(`Email ${email.subject}: ${relevantAttachments.length} relevant attachments found`);

        // Lire le contenu des pièces jointes pertinentes
        const attachmentContents = [];
        for (const attachment of relevantAttachments) {
          const content = await readAttachmentContent(
            {
              ...attachment,
              parentId: email.id,
              messageId: email.id,
            },
            accessToken,
            userEmail
          );

          if (content) {
            attachmentContents.push(content);
          }
        }

        emailContents.push({
          id: email.id,
          subject: email.subject || "Sans objet",
          sender: email.from ? email.from.emailAddress.address : "Expéditeur inconnu",
          senderName: email.from ? email.from.emailAddress.name : "Nom inconnu",
          body: email.body.content || "",
          bodyType: email.body.contentType || "text",
          receivedDateTime: email.receivedDateTime,
          attachments: relevantAttachments,
          attachmentContents: attachmentContents,
        });
      }

      // Trier les emails par date (plus ancien en premier)
      emailContents.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));

      logMessage(`Retrieved ${emailContents.length} emails from conversation with filtered attachments`);

      // Log des pièces jointes pertinentes trouvées
      emailContents.forEach((email, index) => {
        if (email.attachmentContents.length > 0) {
          logMessage(`Email ${index + 1} has ${email.attachmentContents.length} relevant attachments`);
          email.attachmentContents.forEach((att) => {
            logMessage(`- Attachment: ${att.name} (${att.type})`);
          });
        }
      });

      callback(emailContents);
    } catch (error) {
      logMessage(`Error in Graph API call: ${error.message}`);
      getCurrentEmailOnly(callback);
    }
  }

  // Fonction de fallback : récupérer seulement l'email courant
  function getCurrentEmailOnly(callback) {
    logMessage("Using current email only as fallback...");

    try {
      const item = Office.context.mailbox.item;

      // Récupérer le corps de l'email courant
      item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Filtrer les pièces jointes de l'email courant
          const relevantAttachments = item.attachments
            ? item.attachments.filter((att) => isRelevantAttachment(att))
            : [];

          const emailContent = [
            {
              id: item.itemId || "current",
              subject: item.subject || "Sans objet",
              sender: item.from ? item.from.emailAddress : "Expéditeur inconnu",
              senderName: item.from ? item.from.displayName : "Nom inconnu",
              body: result.value,
              bodyType: "html",
              receivedDateTime: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString(),
              attachments: relevantAttachments,
              attachmentContents: relevantAttachments.map((att) => ({
                name: att.name,
                type: att.contentType,
                content: "[Pièce jointe disponible - contenu non accessible en mode fallback]",
              })),
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

  // Fonction pour traiter les pièces jointes filtrées
  function processRelevantAttachments(attachmentContents) {
    if (!attachmentContents || attachmentContents.length === 0) {
      return "Aucune pièce jointe PDF ou DOCX";
    }

    return attachmentContents
      .map((att) => {
        let info = `- ${att.name} (${att.type})`;
        if (att.size) {
          info += ` - ${Math.round(att.size / 1024)} KB`;
        }
        if (att.content && att.content !== "[Document présent]") {
          info += `\n  Contenu: ${att.content}`;
        }
        return info;
      })
      .join("\n");
  }

  // Fonction pour créer un contexte de conversation complet
  function buildConversationContext(emailContents) {
    let context = "=== CONTEXTE DE LA CONVERSATION COMPLÈTE ===\n\n";

    emailContents.forEach((email, index) => {
      context += `--- Email ${index + 1} ---\n`;
      context += `Date: ${new Date(email.receivedDateTime).toLocaleString("fr-FR")}\n`;
      context += `De: ${email.senderName} (${email.sender})\n`;
      context += `Objet: ${email.subject}\n\n`;

      // Nettoyer le contenu HTML basique
      let cleanBody = email.body
        .replace(/<[^>]*>/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      context += `Contenu:\n${cleanBody}\n\n`;

      // Ajouter les informations sur les pièces jointes pertinentes
      if (email.attachmentContents && email.attachmentContents.length > 0) {
        context += `Pièces jointes (PDF/DOCX):\n`;
        context += processRelevantAttachments(email.attachmentContents);
        context += "\n\n";
      }

      context += "---\n\n";
    });

    return context;
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
      max_tokens: 4096,
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

          // Créer le contexte complet de la conversation
          const conversationContext = buildConversationContext(emailContents);

          // Créer un prompt enrichi avec le contexte complet
          const fullPrompt = `${conversationContext}

=== VOTRE QUESTION ===
${prompt}

=== INSTRUCTIONS ===
Veuillez analyser l'ensemble de cette conversation email en tenant compte de tous les messages et des pièces jointes PDF/DOCX mentionnées. 

Répondez à ma question en vous basant sur le contexte complet de la conversation et des documents joints.

En plus de répondre à ma question, pourriez-vous également suggérer un objet approprié pour ma réponse? Présentez-le sous la forme "Objet suggéré: [votre suggestion d'objet]" à la fin de votre réponse.

Veuillez répondre en français et de manière professionnelle.`;

          console.log("Appel de l'API Mistral avec le contexte complet...");
          logMessage(
            `Sending ${emailContents.length} emails and ${emailContents.reduce((total, email) => total + (email.attachmentContents?.length || 0), 0)} relevant attachments to AI`
          );

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
