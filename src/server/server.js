/* eslint-disable no-undef */
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

app.post("/api/callMistralAI", async (req, res) => {
  try {
    const { prompt } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: "Prompt is required" });
    }

    const mistralUrl =
      "https://iebdevias01.services.ai.azure.com/models/chat/completions?api-version=2024-05-01-preview";
    const mistralApiKey = "BYlb7M6GuQut9GkmzTcwWpRzvSp0XdCbU2hNwW2JZfWoNPNPHQtFJQQJ99BEACfhMk5XJ3w3AAAAACOGH4rr";
    const mistralModel = "mistral-large-2411";

    const response = await axios.post(
      mistralUrl,
      {
        model: mistralModel,
        messages: [{ role: "user", content: prompt }],
      },
      {
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${mistralApiKey}`,
        },
      }
    );

    res.json(response.data);
  } catch (error) {
    console.error("Error calling Mistral AI:", error.response?.data || error.message);
    res.status(500).json({
      error: "Failed to call Mistral AI",
      details: error.response?.data || error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
