import axios from "axios";

class GraphHelper {
  constructor() {
    this._tokenPromise = null;
    this._tokenExpiration = null;
    this.delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  }

  logMessage(message) {
    $("#log-area").append(`<div>${new Date().toISOString()}: ${message}</div>`);
  }

  /**
   * Gets application token.
   * @returns {Promise<string>} Application token.
   */
  async getAccessToken() {
    this.logMessage("ho");
    const response = await axios.post("http://localhost:3001/api/getAccessToken");
    this.logMessage("test access");

    return response.access_token;
  }

  /**
   * Make a request to Microsoft Graph API
   * @param {string} method - HTTP method
   * @param {string} endpoint - API endpoint
   * @param {object} [data] - Request data
   * @param {object} [headers] - Additional headers
   * @returns {Promise<any>} Response data
   */
  async graphRequest(method, endpoint, data = null, headers = {}) {
    const token = await this.getAccessToken();
    const config = {
      method: method,
      url: `https://graph.microsoft.com/v1.0${endpoint}`,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...headers,
      },
    };

    if (data) {
      config.data = data;
    }

    try {
      const response = await axios(config);
      return response.data;
    } catch (error) {
      console.error("Graph API request failed:", error);
      throw error;
    }
  }

  /**
   * Saves a DOCX buffer to the SharePoint site associated with a Teams meeting
   * @param {Buffer} docxBuffer - The DOCX file buffer
   * @param {string} userId - The user's AAD Object ID
   * @param {string} fileName - The name for the saved file
   * @returns {Promise<boolean>} True if successful
   */
  async saveDocxToSharePoint(docxBuffer, userId, fileName) {
    const token = await this.getAccessToken();

    try {
      const uploadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/Shared/${fileName}:/content`;
      const response = await axios.put(uploadUrl, docxBuffer, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/octet-stream",
        },
      });
      return response.data.webUrl;
    } catch (error) {
      console.error("Error saving file to SharePoint:", error.message);
      throw new Error("Failed to save file to SharePoint: " + error.message);
    }
  }

  /**
   * Gets the meeting transcript for the passed meeting Id.
   * @param {string} meetingId Id of the meeting
   * @returns Transcript of meeting if any therwise return empty string.
   */
  async GetMeetingTranscriptionsAsync(userId, meetingId) {
    const MAX_ATTEMPTS = 30;
    const RETRY_DELAY = 1000; // 1 second between retries

    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
      try {
        const access_Token = await this.getAccessToken();
        const getAllTranscriptsEndpoint = `${process.env.GraphApiEndpoint}/users/${userId}/onlineMeetings/${meetingId}/transcripts`;

        const transcriptsResponse = await axios({
          method: "get",
          url: getAllTranscriptsEndpoint,
          headers: {
            Authorization: `Bearer ${access_Token}`,
          },
        });

        const transcripts = transcriptsResponse.data.value;

        // If no transcripts, wait and retry
        if (!transcripts || transcripts.length === 0) {
          if (attempt === MAX_ATTEMPTS - 1) {
            throw new Error("No transcripts found after maximum attempts");
          }
          console.log(`No transcripts yet. Attempt ${attempt + 1}/${MAX_ATTEMPTS}. Retrying in ${RETRY_DELAY}ms...`);
          await this.delay(RETRY_DELAY);
          continue;
        }

        // Sort transcripts by creation date in descending order and take the most recent one
        transcripts.sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
        const latestTranscript = transcripts[0];

        // Fetch the content of the most recent transcript
        const transcriptResponse = await axios({
          method: "get",
          url: `${getAllTranscriptsEndpoint}/${latestTranscript.id}/content?$format=text/vtt`,
          headers: {
            Authorization: `Bearer ${access_Token}`,
          },
        });
        return transcriptResponse.data;
      } catch (error) {
        // Immediately fail on 403
        if (error.response?.status === 403) {
          throw new Error("Access forbidden (403) for user :" + userId + " on meeting " + meetingId);
        }

        if (error.response?.status === 402) {
          throw new Error("Payment required (402), exceeded transcript fetch limit");
        }

        // For other errors on last attempt, throw error
        if (attempt === MAX_ATTEMPTS - 1) {
          throw error;
        }

        // For other errors, retry
        console.log(`Error occurred. Attempt ${attempt + 1}/${MAX_ATTEMPTS}. Retrying in ${RETRY_DELAY}ms...`);
        await this.delay(RETRY_DELAY);
      }
    }
  }

  /**
   * Télécharge le contenu d'un fichier, avec support pour les données binaires
   * @param {string} downloadUrl - URL de téléchargement du fichier
   * @param {boolean} [binary=false] - Indiquer si le fichier doit être téléchargé en mode binaire
   * @returns {Promise<string|Buffer>} - Contenu du fichier (texte ou Buffer binaire)
   */
  async DownloadFileContent(downloadUrl, binary = false) {
    try {
      const accessToken = await this.getAccessToken();

      const responseType = binary ? "arraybuffer" : "text";

      const response = await axios({
        method: "get",
        url: downloadUrl,
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        responseType: responseType,
      });

      // Si binaire, retourner un Buffer, sinon la chaîne de caractères
      // eslint-disable-next-line no-undef
      return binary ? Buffer.from(response.data) : response.data;
    } catch (error) {
      console.error("Error downloading file content:", error);
      throw error;
    }
  }
}

export default GraphHelper;
