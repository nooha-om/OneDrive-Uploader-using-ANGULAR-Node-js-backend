import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
app.use(express.json());
app.use(cors());

const USER_ID = "34f820f7-8fed-4f01-a74d-3b5c98a8de86";
const APP_TOKEN = "YOUR_LONG_APP_TOKEN_HERE"; 

// Upload file to OneDrive
app.post("/onedrive/upload", async (req, res) => {
  const { filename, content } = req.body;

  try {
    const graphResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root:/${filename}:/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${APP_TOKEN}`,
          "Content-Type": "text/plain",
        },
        body: content,
      }
    );

    if (!graphResponse.ok) {
      const errorText = await graphResponse.text();
      throw new Error(errorText);
    }

    const data = await graphResponse.json();
    res.json({ fileId: data.id, message: "File uploaded successfully" });
  } catch (err) {
    console.error("File upload failed:", err);
    res.status(500).send("File upload failed");
  }
});

// Load file by name
app.get("/onedrive/fileByName/:filename", async (req, res) => {
  const filename = req.params.filename;

  try {
    const searchResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root/search(q='${filename}')`,
      { headers: { Authorization: `Bearer ${APP_TOKEN}` } }
    );

    const searchData = await searchResponse.json();
    if (!searchData.value || searchData.value.length === 0) {
      return res.status(404).json({ error: "File not found", searchedFor: filename });
    }

    const fileId = searchData.value[0].id;

    const contentResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/items/${fileId}/content`,
      { headers: { Authorization: `Bearer ${APP_TOKEN}` } }
    );

    const content = await contentResponse.text();
    res.json({ fileId, filename, content });
  } catch (err) {
    console.error("File fetch failed:", err);
    res.status(500).send("File fetch failed");
  }
});

// Update file by name
app.put("/onedrive/fileByName/:filename", async (req, res) => {
  const filename = req.params.filename;
  const { content } = req.body;

  try {
    const searchResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root/search(q='${filename}')`,
      { headers: { Authorization: `Bearer ${APP_TOKEN}` } }
    );

    const searchData = await searchResponse.json();
    if (!searchData.value || searchData.value.length === 0) {
      return res.status(404).json({ error: "File not found", searchedFor: filename });
    }

    const fileId = searchData.value[0].id;

    const updateResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/items/${fileId}/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${APP_TOKEN}`,
          "Content-Type": "text/plain",
        },
        body: content,
      }
    );

    if (!updateResponse.ok) {
      const errorText = await updateResponse.text();
      throw new Error(errorText);
    }

    res.json({ fileId, filename, message: "File updated successfully" });
  } catch (err) {
    console.error("File update failed:", err);
    res.status(500).send("File update failed");
  }
});

// Spellcheck API (LanguageTool)
app.post("/spellcheck", async (req, res) => {
  const { text } = req.body;

  try {
    const response = await fetch("https://api.languagetool.org/v2/check", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: `text=${encodeURIComponent(text)}&language=en-US`
    });

    const data = await response.json();
    res.json(data);
  } catch (err) {
    console.error("Spell check failed:", err);
    res.status(500).json({ error: "Spell check failed" });
  }
});

//  Load + Spellcheck file
app.get("/onedrive/checkFile/:filename", async (req, res) => {
  const filename = req.params.filename;

  try {
    const searchResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root/search(q='${filename}')`,
      { headers: { Authorization: `Bearer ${APP_TOKEN}` } }
    );
    const searchData = await searchResponse.json();

    if (!searchData.value || searchData.value.length === 0) {
      return res.status(404).json({ error: "File not found", searchedFor: filename });
    }

    const fileId = searchData.value[0].id;

    const contentResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/items/${fileId}/content`,
      { headers: { Authorization: `Bearer ${APP_TOKEN}` } }
    );
    const content = await contentResponse.text();

    const spellcheckResponse = await fetch("https://api.languagetool.org/v2/check", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: `text=${encodeURIComponent(content)}&language=en-US`
    });
    const spellData = await spellcheckResponse.json();

    const errors = spellData.matches.map(m => ({
      word: content.substring(m.offset, m.offset + m.length),
      suggestions: m.replacements.map(r => r.value)
    }));

    res.json({ fileId, filename, content, errors });
  } catch (err) {
    console.error("File check failed:", err);
    res.status(500).send("File check failed");
  }
});

app.listen(3000, () => console.log("Backend running on http://localhost:3000"));
