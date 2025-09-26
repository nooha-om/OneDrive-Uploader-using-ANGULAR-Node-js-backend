import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
app.use(express.json());
app.use(cors());

const USER_ID = "34f820f7-8fed-4f01-a74d-3b5c98a8de86";
const APP_TOKEN = "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImYzSG5CT1JxMU9aXzR4NWJoc2VGR0w2aXJUcEhkQWpUUTNkeDRDbnJhZkkiLCJhbGciOiJSUzI1NiIsIng1dCI6IkhTMjNiN0RvN1RjYVUxUm9MSHdwSXEyNFZZZyIsImtpZCI6IkhTMjNiN0RvN1RjYVUxUm9MSHdwSXEyNFZZZyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80NDVlNGI3MC1lMzg2LTRkNzMtYmIzMC01MThmNWUxYWMxYWUvIiwiaWF0IjoxNzU4ODc5MzgxLCJuYmYiOjE3NTg4NzkzODEsImV4cCI6MTc1ODg4MzI4MSwiYWlvIjoiazJKZ1lPalQ3ejVaSXhMbUtLbWNkdlhrNXYzaEFBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJkeWF6IiwiYXBwaWQiOiI4ZWU3YmViMy03YmZjLTRhODctOWE1OS1mMjMxN2YzNjI5MTUiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80NDVlNGI3MC1lMzg2LTRkNzMtYmIzMC01MThmNWUxYWMxYWUvIiwiaWR0eXAiOiJhcHAiLCJvaWQiOiIxZmEzNmI2Ni1kMDRjLTQ5NTYtOTZmOC1kZmRiODg5YWI1NzMiLCJyaCI6IjEuQVZZQWNFdGVSSWJqYzAyN01GR1BYaHJCcmdNQUFBQUFBQUFBd0FBQUFBQUFBQUNmQUFCV0FBLiIsInJvbGVzIjpbIkRpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIiwiU2l0ZXMuUmVhZC5BbGwiLCJTaXRlcy5SZWFkV3JpdGUuQWxsIiwiR3JvdXAuUmVhZFdyaXRlLkFsbCIsIkZpbGVzLlJlYWRXcml0ZS5BbGwiLCJEaXJlY3RvcnkuUmVhZC5BbGwiXSwic3ViIjoiMWZhMzZiNjYtZDA0Yy00OTU2LTk2ZjgtZGZkYjg4OWFiNTczIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiNDQ1ZTRiNzAtZTM4Ni00ZDczLWJiMzAtNTE4ZjVlMWFjMWFlIiwidXRpIjoidUU2WW1kMVJxRU9rdVpiZkdWOG1BQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMDk5N2ExZDAtMGQxZC00YWNiLWI0MDgtZDVjYTczMTIxZTkwIl0sInhtc19mdGQiOiJXNnZpV2FnTVV5RS11d0hUdHRPVlNPYktBVllhYk92TWpkaGdnUXl6anZJQmEyOXlaV0Z6YjNWMGFDMWtjMjF6IiwieG1zX2lkcmVsIjoiNyAxNCIsInhtc19yZCI6IjAuNDJMbFlCSmlEQk1TNFdBWEV0anJWSzZnd05EczBGRWVfVF9yMHRVdW9DaW5rSUNTTE9lQi1WOFgtdTA0SzFyWWNvajdERkNVUTBpQW1RRUNEa0JwQUEiLCJ4bXNfdGNkdCI6MTY3MjIwMjQyNH0.eMWoZwE8VKoq6xBg8y2SQ2C5ezHs3D35B-hM5g_Q25jGIN6Gg8UzEtapHUFODYpuQJ1FEGA_8MtuIZkDKGSgCQ010deMW8I5V_bcQcL2IGhCPe0KgdOgV25aBoJvkUmT9RZPM9DeGNNDC4aJ2yVehmAeJhRTEKl3mkZJ456hTNc6CS19LodTJ8sVnrr8uXXyD2sAaYtAyu9XvMknBrDiHoMdDPfnL8HS763NePXn3MNQmxBYa1Cw0Fm_tBNsGvf2QD9V21l-l4RI5ARVpfW-c6b-gjkLUgU0ivnYnWj3Ns6QN-ie615uAnoSe60wEFkCB5ce-O5ujXpd8Np236eS1g"; 

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

//  API (LanguageTool)
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

app.listen(3000, () => console.log("Backend running on http://localhost:3000"));
