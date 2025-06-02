import axios from "axios";

class GraphHelper {
  constructor() {
    this._tokenPromise = null;
    this._tokenExpiration = null;
  }

  /**
   * Gets application token.
   * @returns {Promise<string>} Application token.
   */
  async getAccessToken() {
    const response = await axios.post("http://localhost:3001/api/getAccessToken");
    return response.data.access_token;
  }
}

export default GraphHelper;
