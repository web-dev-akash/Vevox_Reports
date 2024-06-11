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
const { pLimit } = require("plimit-lit");
require("dotenv").config();

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(cors());

const limit = pLimit(20);

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

const findUserWithEmail = async (email, config) => {
  try {
    const contact = await axios.get(
      `https://www.zohoapis.com/crm/v2/Contacts/search?email=${email}`,
      config
    );

    if (contact.status >= 400 || contact.status === 204) {
      return {
        status: contact.status,
      };
    }

    return {
      status: 200,
      id: contact.data.data[0].id,
    };
  } catch (error) {
    return {
      status: 500,
      message: error.message,
      email: email,
    };
  }
};

const getAnalysisData = async (query, zohoConfig) => {
  try {
    const response = await axios.post(
      `https://www.zohoapis.com/crm/v3/coql`,
      { select_query: query },
      zohoConfig
    );
    if (response.status >= 400 || response.status === 204) {
      return {
        status: response.status,
      };
    }
    return {
      status: response.status,
      id: response.data.data[0].id,
      sessionDate: response.data.data[0].Session_Date_Time,
    };
  } catch (error) {
    throw error;
  }
};

const addAttemptToZoho = async (body, config) => {
  try {
    const attemptsres = await axios.post(
      `https://www.zohoapis.com/crm/v3/Attempts`,
      body,
      config
    );
    return {
      status: attemptsres.status,
    };
  } catch (error) {
    return {
      status: 500,
      message: error.message,
    };
  }
};

const updateDataonZoho = async (users) => {
  const token = await getZohoToken();
  const config = {
    headers: {
      Authorization: `Zoho-oauthtoken ${token}`,
      "Content-Type": "application/json",
    },
  };

  const attemptsData = await Promise.all(
    users.map(async (user) => {
      const sessionQuery = `select id, Session_Date_Time from Sessions where Vevox_Session_ID = '${user.sessionId}'`;
      const [contact, session] = await Promise.all([
        limit(() => findUserWithEmail(user.email, config)),
        limit(() => getAnalysisData(sessionQuery, config)),
      ]);
      if (contact.status === 200 && session.status === 200) {
        return {
          contactId: contact.id,
          sessionId: session.id,
          score: user.correct,
          sessionDate: session.sessionDate,
        };
      }
      return;
    })
  );

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

  const sendData = async () => {
    const addAttempts = await addAttemptToZoho(body, config);
    body.data = [];
  };

  for (const attempt of attemptsData) {
    const attemptsQuery = `select Contact_Name.id as contactId from Attempts where Contact_Name = '${attempt.contactId}' and Session = '${attempt.sessionId}'`;
    const [attempts] = await Promise.all([
      limit(() => getAnalysisData(attemptsQuery, config)),
    ]);
    if (attempts.status !== 200) {
      attemptNumber += 1;
      body.data.push({
        Name: `${attemptNumber}`,
        Contact_Name: attempt.contactId,
        Session: attempt.sessionId,
        Quiz_Score: attempt.score,
        Session_Date_Time: attempt.sessionDate,
      });
      if (body.data.length === 100) {
        await sendData();
      }
    }
  }
  if (body.data.length > 0) {
    await sendData();
  }
  return { result: "Success" };
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

function getLastPollNumber(data) {
  const keys = Object.keys(data);
  const emptyKeys = keys.filter((key) => key.startsWith("__EMPTY_"));
  const lastKey = emptyKeys
    .sort((a, b) => {
      const numA = parseInt(a.split("_")[1]);
      const numB = parseInt(b.split("_")[1]);
      return numA - numB;
    })
    .pop();

  return data[lastKey];
}

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
        const email = data1[i]["__EMPTY_1"].trim();
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
      const totalPolled = getLastPollNumber(data2[0]);
      for (let i = 7; i < data2.length; i++) {
        const firstname = data2[i]["Polling Results"];
        const lastname = data2[i]["__EMPTY"];
        const email = data2[i]["__EMPTY_1"].trim();
        if (!firstname && !email) {
          continue;
        }
        const correct = data2[i]["__EMPTY_2"] ? data2[i]["__EMPTY_2"] : 0;
        const totalNotEmpty = Object.values(data2[i]).filter(
          (val) => val !== ""
        ).length;

        const totalAttempted =
          totalPolled - (Object.keys(data2[i]).length - totalNotEmpty);

        const userFound = currentUsers.find((user) => user.email === email);

        if (userFound) {
          const obj = {
            firstname,
            lastname,
            correct,
            email: email,
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
    return res.status(200).send(data);
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
