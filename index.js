const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { google } = require("googleapis");
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const { promisify } = require("util");
const unlinkAsync = promisify(fs.unlink);
require("dotenv").config();
const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(cors());

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const PORT = process.env.PORT || 8080;

const storage = multer.diskStorage({
  destination: path.join(__dirname, "uploads"),
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, "template")));

app.get("/", (req, res) => {
  res.sendFile(`index.html`);
});

const getZohoToken = async () => {
  try {
    const res = await axios.post(
      `https://accounts.zoho.com/oauth/v2/token?client_id=${CLIENT_ID}&grant_type=refresh_token&client_secret=${CLIENT_SECRET}&refresh_token=${REFRESH_TOKEN}`
    );
    console.log(res.data);
    const token = res.data.access_token;
    return token;
  } catch (error) {
    res.send({
      error,
    });
  }
};

const getVevoxSheetData = async () => {
  const spreadsheetId = process.env.SPREADSHEET_ID;
  const auth = new google.auth.GoogleAuth({
    keyFile: "key.json", //the key file
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const authClientObject = await auth.getClient();
  const sheet = google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  const readData = await sheet.spreadsheets.values.get({
    auth, //auth object
    spreadsheetId, // spreadsheet id
    range: "Vevox Data!A:H", //range of cells to read from.
  });
  return readData.data.values;
};

const addDataToSheet = async (users) => {
  const spreadsheetId = process.env.SPREADSHEET_ID;
  const auth = new google.auth.GoogleAuth({
    keyFile: "key.json", //the key file
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const authClientObject = await auth.getClient();
  const sheet = google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  const rowsLength = await getVevoxSheetData();
  const writeData = await sheet.spreadsheets.values.update({
    auth, //auth object
    spreadsheetId, //spreadsheet id
    range: `Vevox Data!A${rowsLength.length + 1}:H${
      rowsLength.length + users.length
    }`, //sheet name and range of cells
    valueInputOption: "USER_ENTERED",
    resource: {
      values: users,
    },
  });
  return writeData.data;
};

const updateDataonZoho = async (users) => {
  const token = await getZohoToken();
  const config = {
    headers: {
      Authorization: `Zoho-oauthtoken ${token}`,
      "Content-Type": "application/json",
    },
  };

  const attemptsData = [];
  const playersData = [];
  for (let i = 0; i < users.length; i++) {
    const contact = await axios.get(
      `https://www.zohoapis.com/crm/v2/Contacts/search?email=${users[i].email}`,
      config
    );
    if (contact.status >= 400 || contact.status === 204) {
      continue;
    }

    const contactId = contact.data.data[0].id;
    const name = contact.data.data[0].Student_Name;
    const email = users[i].email;
    const score = users[i].correct;

    playersData.push({
      player_name: name,
      player_email: email,
      player_external_id: contactId,
      offline: "1",
    });

    const session = await axios.get(
      `https://www.zohoapis.com/crm/v2/Sessions/search?criteria=((Vevox_Session_ID:equals:${users[i].sessionId}))`,
      config
    );
    if (session.status >= 400 || session.status === 204) {
      continue;
    }
    const totalSessions = session.data.data[0];
    const sessionId = totalSessions.id;
    attemptsData.push({
      contactId,
      sessionId,
      score: score,
      sessionDate: totalSessions.Session_Date_Time,
    });
  }

  const attemptsCount = await axios.get(
    `https://www.zohoapis.com/crm/v2.1/Attempts/actions/count`,
    config
  );

  let attemptNumber = attemptsCount.data.count;
  const body = {
    data: [],
    apply_feature_execution: [
      {
        name: "layout_rules",
      },
    ],
    trigger: ["workflow"],
  };
  for (let i = 0; i < attemptsData.length; i++) {
    const attempts = await axios.get(
      `https://www.zohoapis.com/crm/v2/Attempts/search?criteria=((Contact_Name:equals:${attemptsData[i].contactId})and(Session:equals:${attemptsData[i].sessionId}))`,
      config
    );
    if (attempts.status === 200) {
      console.log("Attempts Already Exists");
      continue;
    }
    attemptNumber = attemptNumber + 1;
    body.data.push({
      Name: `${attemptNumber}`,
      Contact_Name: attemptsData[i].contactId,
      Session: attemptsData[i].sessionId,
      Quiz_Score: attemptsData[i].score,
      Session_Date_Time: attemptsData[i].sessionDate,
      $append_values: {
        Name: true,
        Contact_Name: true,
        Session: true,
        Quiz_Score: true,
        Session_Date_Time: true,
      },
    });
  }

  if (body.data.length > 0) {
    const attemptsres = await axios.post(
      `https://www.zohoapis.com/crm/v3/Attempts/upsert`,
      body,
      config
    );
    console.log("Zoho Attempts Updated");
  }

  const api_key = process.env.POINTAGRAM_API_KEY;
  const api_user = process.env.POINTAGRAM_API_USER;
  const pointagramConfig = {
    headers: {
      api_key: api_key,
      "Content-Type": "application/json",
      api_user: api_user,
    },
  };
  const requests = playersData.map(async (player) => {
    const email = player.player_email;
    const listPlayers = await axios.get(
      `https://app.pointagram.com/server/externalapi.php/list_players?search_by=Email&filter=${email}`,
      pointagramConfig
    );

    if (listPlayers.data?.length === 0) {
      console.log("Creating a new player:", email);
      await axios.post(
        `https://app.pointagram.com/server/externalapi.php/create_player`,
        player,
        pointagramConfig
      );
    }
  });
  await Promise.all(requests);
  return { message: "SUCCESS" };
};

const updateDataOnVevoxSheet = async (users) => {
  const vevoxData = await getVevoxSheetData();
  const addVevoxData = [];
  for (let i = 0; i < users.length; i++) {
    const foundExistingEntry = vevoxData.filter(
      (user) => user[3] === users[i].email && +user[5] === +users[i].sessionId
    );
    if (foundExistingEntry.length == 0) {
      addVevoxData.push([
        users[i].firstname,
        users[i].lastname,
        users[i].correct,
        users[i].email,
        users[i].date,
        users[i].sessionId,
        users[i].attempted,
        users[i].polled,
      ]);
    }
  }
  await addDataToSheet(addVevoxData);
  return addVevoxData;
};

app.post("/view", upload.array("file", 50), async (req, res) => {
  const files = req.files;
  if (files.length === 0) {
    return res
      .status(400)
      .send(
        `<h1 style="display:grid;place-items:center;min-height:100vh;">No files were uploaded.</h1>`
      );
  }
  try {
    const finalUsers = [];
    for (const file of files) {
      console.log(file.path);
      const workbook = xlsx.readFile(file.path);
      const sheetName1 = workbook.SheetNames[0];
      const sheet1 = workbook.Sheets[sheetName1];
      const data1 = xlsx.utils.sheet_to_json(sheet1);
      const sessionId = data1[4]["__EMPTY"];
      const currentUsers = [];
      for (let i = 8; i < data1.length; i++) {
        const firstname = data1[i][""];
        const lastname = data1[i]["__EMPTY"];
        const email = data1[i]["__EMPTY_1"];
        const attemptDate = new Date(
          data1[i]["__EMPTY_2"].substring(0, 11)
        ).toDateString();
        const obj = { firstname, lastname, email, attemptDate };
        if (email && !email.includes("1234500")) {
          currentUsers.push(obj);
        }
      }

      const sheetName2 = workbook.SheetNames[2];
      const sheet2 = workbook.Sheets[sheetName2];
      const data2 = xlsx.utils.sheet_to_json(sheet2);
      const totalPolled = Object.values(data2[2]).length - 3;
      for (let i = 7; i < data2.length - 2; i++) {
        const firstname = data2[i]["Polling Results"];
        const lastname = data2[i]["__EMPTY"];
        const correct = data2[i]["__EMPTY_1"] ? data2[i]["__EMPTY_1"] : 0;
        const totalNotEmpty = Object.values(data2[i]).filter(
          (val) => val === ""
        ).length;

        const totalAttempted =
          totalNotEmpty === 0 ? 0 : totalPolled - totalNotEmpty;
        const userFound = currentUsers.find(
          (user) => user.firstname === firstname && user.lastname === lastname
        );
        if (userFound) {
          const obj = {
            firstname,
            lastname,
            correct,
            email: userFound.email,
            date: userFound.attemptDate,
            polled: totalPolled,
            attempted: totalAttempted,
            sessionId,
          };
          finalUsers.push(obj);
        }
      }
    }
    for (const file of files) {
      await unlinkAsync(file.path);
    }
    await updateDataOnVevoxSheet(finalUsers);
    const data = await updateDataonZoho(finalUsers);
    return res.status(200).send({ data: data });
  } catch (error) {
    console.error("Error reading Excel file:", error);
    for (const file of files) {
      await unlinkAsync(file.path);
    }
    return res.status(500).send({ error: "Error reading Excel file." });
  }
});

app.listen(PORT, () => {
  console.log(`http://localhost:${PORT}`);
});
