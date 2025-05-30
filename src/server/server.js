import express, { json } from "express";
import axios from "axios";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3001;

app.use(cors());
app.use(json());

// Endpoint original avec URLSearchParams
app.post("/api/getAccessToken", async (req, res) => {
  try {
    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", "d02b12a0-cb0b-48d7-8c33-0ce1952eae0e");
    params.append("client_secret", "6c28Q~fEQWB9F8chLy.FMcr5a8bWxAWPtgjCfct2");
    params.append("scope", "https://graph.microsoft.com/.default");

    const response = await axios.post(
      "https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token",
      params.toString(),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );

    res.json(response.data);
  } catch (error) {
    console.error("Error retrieving access token:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to get access token",
      details: error.response?.data,
    });
  }
});

// Nouveau endpoint qui reproduit exactement le curl
app.post("/api/getAccessTokenWithCookies", async (req, res) => {
  try {
    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", "d02b12a0-cb0b-48d7-8c33-0ce1952eae0e");
    params.append("client_secret", "6c28Q~fEQWB9F8chLy.FMcr5a8bWxAWPtgjCfct2");
    params.append("scope", "https://graph.microsoft.com/.default");

    const response = await axios.post(
      "https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token",
      params.toString(),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          Cookie:
            "fpc=Atg5Mfieyz9Pk1WcIFM246QPKDc-AQAAALeTy98OAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd",
        },
      }
    );

    res.json(response.data);
  } catch (error) {
    console.error("Error retrieving access token with cookies:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to get access token with cookies",
      details: error.response?.data,
    });
  }
});

// Endpoint pour comparer les deux tokens
app.post("/api/compareTokens", async (req, res) => {
  try {
    // Token sans cookies
    const params1 = new URLSearchParams();
    params1.append("grant_type", "client_credentials");
    params1.append("client_id", "d02b12a0-cb0b-48d7-8c33-0ce1952eae0e");
    params1.append("client_secret", "6c28Q~fEQWB9F8chLy.FMcr5a8bWxAWPtgjCfct2");
    params1.append("scope", "https://graph.microsoft.com/.default");

    const response1 = await axios.post(
      "https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token",
      params1.toString(),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      }
    );

    // Token avec cookies
    const params2 = new URLSearchParams();
    params2.append("grant_type", "client_credentials");
    params2.append("client_id", "d02b12a0-cb0b-48d7-8c33-0ce1952eae0e");
    params2.append("client_secret", "6c28Q~fEQWB9F8chLy.FMcr5a8bWxAWPtgjCfct2");
    params2.append("scope", "https://graph.microsoft.com/.default");

    const response2 = await axios.post(
      "https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token",
      params2.toString(),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          Cookie:
            "fpc=Atg5Mfieyz9Pk1WcIFM246QPKDc-AQAAALeTy98OAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd",
        },
      }
    );

    // Décoder les tokens pour comparaison
    const decodeToken = (token) => {
      const parts = token.split(".");
      return JSON.parse(Buffer.from(parts[1], "base64").toString());
    };

    const token1Decoded = decodeToken(response1.data.access_token);
    const token2Decoded = decodeToken(response2.data.access_token);

    res.json({
      tokenWithoutCookies: {
        token: response1.data.access_token.substring(0, 50) + "...",
        decoded: token1Decoded,
        identical: response1.data.access_token === response2.data.access_token,
      },
      tokenWithCookies: {
        token: response2.data.access_token.substring(0, 50) + "...",
        decoded: token2Decoded,
        identical: response1.data.access_token === response2.data.access_token,
      },
      tokensAreIdentical: response1.data.access_token === response2.data.access_token,
    });
  } catch (error) {
    console.error("Error comparing tokens:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to compare tokens",
      details: error.response?.data,
    });
  }
});

// Endpoint pour tester l'API Graph avec le token généré
app.post("/api/testGraphAPI", async (req, res) => {
  try {
    // Générer le token
    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", "d02b12a0-cb0b-48d7-8c33-0ce1952eae0e");
    params.append("client_secret", "6c28Q~fEQWB9F8chLy.FMcr5a8bWxAWPtgjCfct2");
    params.append("scope", "https://graph.microsoft.com/.default");

    const tokenResponse = await axios.post(
      "https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token",
      params.toString(),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      }
    );

    const token = tokenResponse.data.access_token;

    // Tester l'API Graph
    const graphResponse = await axios.get("https://graph.microsoft.com/v1.0/users", {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    res.json({
      success: true,
      userCount: graphResponse.data.value?.length || 0,
      tokenPreview: token.substring(0, 50) + "...",
    });
  } catch (error) {
    console.error("Error testing Graph API:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to test Graph API",
      details: error.response?.data || error.message,
      status: error.response?.status,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
