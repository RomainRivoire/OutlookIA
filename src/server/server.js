import express, { json } from "express";
import axios from "axios";
import { stringify } from "qs";
import cors from "cors";

const app = express();
const PORT = 3001;

// Enable CORS for all routes
app.use(cors());

app.use(json());

app.post("/api/getAccessToken", async (req, res) => {
  console.log("test back : ");
  
  try {
    const data = stringify({
      grant_type: "client_credentials",
      client_id: "c4168176-c3b7-4c26-b9a7-ee28cc22faf3",
      scope: "https://graph.microsoft.com/.default",
      client_secret: "cgk8Q~lgnmR.xgnT3Epz58cU~ej20Jtx2CIU9bJx",
    });

    const config = {
      method: "post",
      url: `https://login.microsoftonline.com/bb948182-0852-4388-a948-98f0c3dfb67f/oauth2/v2.0/token`,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      data: data,
    };

    const response = await axios(config);
    console.log("Réponse :", JSON.stringify(response.data, null, 2));
    res.json(response.data);
  } catch (error) {
    console.log(error.message);
    res.status(500).json({ error: "Failed to get access token" });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});