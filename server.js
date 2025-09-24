import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
app.use(express.json());
app.use(cors());


const USER_ID = "34f820f7-8fed-4f01-a74d-3b5c98a8de86";


const APP_TOKEN = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IktNeU5UVjcyVkRlUDMzZmVpTTQxU21TZDMwZnRsTWdSR0ZIWnJyTC1mRW8iLCJhbGciOiJSUzI1NiIsIng1dCI6IkhTMjNiN0RvN1RjYVUxUm9MSHdwSXEyNFZZZyIsImtpZCI6IkhTMjNiN0RvN1RjYVUxUm9MSHdwSXEyNFZZZyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80NDVlNGI3MC1lMzg2LTRkNzMtYmIzMC01MThmNWUxYWMxYWUvIiwiaWF0IjoxNzU4NzE5MjUxLCJuYmYiOjE3NTg3MTkyNTEsImV4cCI6MTc1ODcyMzE1MSwiYWlvIjoiazJSZ1lOQ1dkaEV1dWY2NmVySmhYbWhTVmVNeEFBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJkeWF6IiwiYXBwaWQiOiI4ZWU3YmViMy03YmZjLTRhODctOWE1OS1mMjMxN2YzNjI5MTUiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80NDVlNGI3MC1lMzg2LTRkNzMtYmIzMC01MThmNWUxYWMxYWUvIiwiaWR0eXAiOiJhcHAiLCJvaWQiOiIxZmEzNmI2Ni1kMDRjLTQ5NTYtOTZmOC1kZmRiODg5YWI1NzMiLCJyaCI6IjEuQVZZQWNFdGVSSWJqYzAyN01GR1BYaHJCcmdNQUFBQUFBQUFBd0FBQUFBQUFBQUNmQUFCV0FBLiIsInJvbGVzIjpbIkRpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIiwiU2l0ZXMuUmVhZC5BbGwiLCJTaXRlcy5SZWFkV3JpdGUuQWxsIiwiR3JvdXAuUmVhZFdyaXRlLkFsbCIsIkZpbGVzLlJlYWRXcml0ZS5BbGwiLCJEaXJlY3RvcnkuUmVhZC5BbGwiXSwic3ViIjoiMWZhMzZiNjYtZDA0Yy00OTU2LTk2ZjgtZGZkYjg4OWFiNTczIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiNDQ1ZTRiNzAtZTM4Ni00ZDczLWJiMzAtNTE4ZjVlMWFjMWFlIiwidXRpIjoiSGNRNXU1ZWREMFNWeXFITTNDd05BQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMDk5N2ExZDAtMGQxZC00YWNiLWI0MDgtZDVjYTczMTIxZTkwIl0sInhtc19mdGQiOiJHOERaRVpXcmNPVURSRjhGcXhYZEFNWmtScnBUckhKY0ZvUG9GWjVUNDdRQmEyOXlaV0Z6YjNWMGFDMWtjMjF6IiwieG1zX2lkcmVsIjoiMjIgNyIsInhtc19yZCI6IjAuNDJMbFlCSmlEQk1TNFdBWEV0anJWSzZnd05EczBGRWVfVF9yMHRVdW9DaW5rSUNTTE9lQi1WOFgtdTA0SzFyWWNvajdERkNVUTBpQW1RRUNEa0JwQUEiLCJ4bXNfdGNkdCI6MTY3MjIwMjQyNH0.J7lyj9av7jYaHAw_ZSNcif838QSTVeFI4UspKWdWkPpOatRm6VSne-UVlYJ5XGx_u70G5-evUZ9Y3UjiTehXF6tBO5BQzlbj1XQWsDOJoaenT_Sv4BJBJEhzI0cypoZzAPe0Rv_bKuaPnqo2Ys_beja5O2QBuV2373HvSQ_5rWeH_5aB14DUtcrRa9uv_tkED5vfeuDZe_s3kPRQKhGdrLp3wy5kgK3DBJ-40yYH9yvyhwe59T0uNAc42fLVsgoJGLQFhf5QHmyArcUARQSdtkgIDyZ05m_DoqBdCOJpxW9hsLikPwg1usZP72EmjVzzYkW92KPOhjEu5zK62GWxgw"


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

app.listen(3000, () => console.log("Backend running on http://localhost:3000"));
