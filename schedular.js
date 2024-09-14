const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const bodyParser = require('body-parser');
require('dotenv').config();
const nodemailer = require('nodemailer');
const cron = require('node-cron');
const dayjs = require('dayjs');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3001;

// File path to store email credentials
const emailCredentialsFilePath = path.join(__dirname, 'emailCredentials.json');

// File path to store scheduled email tasks
const scheduledEmailsFilePath = path.join(__dirname, 'scheduledEmails.json');

// CORS configuration
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST','PUT','DELETE','OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(bodyParser.json({ limit: '100mb' }));
app.use(bodyParser.urlencoded({ limit: '100mb', extended: true }));

// Google Sheets and Drive credentials
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
const drive = google.drive({ version: 'v3', auth: oauth2Client });

// Route to view scheduled email tasks
app.get('/scheduled-tasks', (req, res) => {
  try {
    if (!fs.existsSync(scheduledEmailsFilePath)) {
      return res.json({ tasks: [] });
    }

    const tasks = JSON.parse(fs.readFileSync(scheduledEmailsFilePath, 'utf8'));
    const taskSummaries = tasks.map((task, index) => ({
      index,
      emailSubject: task.emailSubject,
      scheduledDateTime: task.scheduledDateTime
    }));

    res.json({ tasks: taskSummaries });
  } catch (error) {
    console.error('Error fetching scheduled tasks:', error);
    res.status(500).json({ message: 'Failed to fetch scheduled tasks' });
  }
});

// Route to delete a scheduled email task by index
app.delete('/delete-scheduled-task/:index', (req, res) => {
  const { index } = req.params;

  try {
    // Read existing tasks
    let tasks = [];
    if (fs.existsSync(scheduledEmailsFilePath)) {
      tasks = JSON.parse(fs.readFileSync(scheduledEmailsFilePath, 'utf8'));
    }

    // Check if the index is valid
    if (index < 0 || index >= tasks.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Remove the task at the specified index
    tasks.splice(index, 1);

    // Save the updated tasks
    fs.writeFileSync(scheduledEmailsFilePath, JSON.stringify(tasks, null, 2), 'utf8');

    res.json({ status: 'success', message: 'Scheduled task deleted successfully.' });
  } catch (error) {
    console.error('Error deleting scheduled task:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while deleting the scheduled task.' });
  }
});
// Route to edit email credentials by index or email
app.put('/edit-email-credentials/:index', (req, res) => {
  const { user, pass } = req.body;
  const { index } = req.params;

  if (!user || !pass) {
    return res.status(400).json({ status: 'error', message: 'Email and password are required.' });
  }

  try {
    // Read existing credentials
    let emailCredentials = [];
    if (fs.existsSync(emailCredentialsFilePath)) {
      emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
    }

    // Check if the index is valid
    if (index < 0 || index >= emailCredentials.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Update credentials at the given index
    emailCredentials[index] = { user, pass };

    // Save the updated credentials
    fs.writeFileSync(emailCredentialsFilePath, JSON.stringify(emailCredentials, null, 2), 'utf8');

    res.json({ status: 'success', message: 'Email credentials updated successfully.' });
  } catch (error) {
    console.error('Error editing email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while editing credentials.' });
  }
});

// Route to delete email credentials by index or email
app.delete('/delete-email-credentials/:index', (req, res) => {
  const { index } = req.params;

  try {
    // Read existing credentials
    let emailCredentials = [];
    if (fs.existsSync(emailCredentialsFilePath)) {
      emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
    }

    // Check if the index is valid
    if (index < 0 || index >= emailCredentials.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Remove the credentials at the specified index
    emailCredentials.splice(index, 1);

    // Save the updated credentials
    fs.writeFileSync(emailCredentialsFilePath, JSON.stringify(emailCredentials, null, 2), 'utf8');

    res.json({ status: 'success', message: 'Email credentials deleted successfully.' });
  } catch (error) {
    console.error('Error deleting email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while deleting credentials.' });
  }
});
// Route to add email credentials
app.post('/add-email-credentials', (req, res) => {
  const { user, pass } = req.body;

  if (!user || !pass) {
    return res.status(400).json({ status: 'error', message: 'Email and password are required.' });
  }

  // Read existing credentials
  let emailCredentials = [];
  if (fs.existsSync(emailCredentialsFilePath)) {
    emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
  }

  // Add new credentials
  emailCredentials.push({ user, pass });

  // Save credentials to file
  fs.writeFileSync(emailCredentialsFilePath, JSON.stringify(emailCredentials, null, 2), 'utf8');

  res.json({ status: 'success', message: 'Email credentials added successfully.' });
});
// Route to get email addresses
app.get('/get-emails', (req, res) => {
  try {
    if (!fs.existsSync(emailCredentialsFilePath)) {
      return res.json({ emails: [] });
    }

    const emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
    const emailAddresses = emailCredentials.map(account => account.user);

    res.json({ emails: emailAddresses });
  } catch (error) {
    console.error('Error fetching email accounts:', error);
    res.status(500).json({ message: 'Failed to fetch email accounts' });
  }
});

/// Route to schedule emails
app.post('/schedule-emails', async (req, res) => {
  const { sheetId, sheetName, emailSubject, emailBody, attachment, ranges, scheduledDateTime } = req.body;

  if (!sheetId || !sheetName || !emailSubject || !emailBody || !attachment || !ranges || !scheduledDateTime) {
    return res.status(400).json({ status: 'error', message: 'One or more parameters are missing.' });
  }

  try {
    // Schedule the emails using cron at the provided scheduledDateTime
    const scheduledDate = dayjs(scheduledDateTime);
    const now = dayjs();

    if (scheduledDate.isBefore(now)) {
      return res.status(400).json({ status: 'error', message: 'Scheduled time must be in the future.' });
    }

    // Convert scheduledDate to cron format
    const cronTime = `${scheduledDate.second()} ${scheduledDate.minute()} ${scheduledDate.hour()} ${scheduledDate.date()} ${scheduledDate.month() + 1} *`;

    // Load existing tasks
    let tasks = [];
    if (fs.existsSync(scheduledEmailsFilePath)) {
      tasks = JSON.parse(fs.readFileSync(scheduledEmailsFilePath, 'utf8'));
    }

    // Add new task
    const task = {
      sheetId,
      sheetName,
      emailSubject,
      emailBody,
      attachment,
      ranges,
      scheduledDateTime,
      cronTime
    };

    tasks.push(task);

    // Save updated tasks
    fs.writeFileSync(scheduledEmailsFilePath, JSON.stringify(tasks, null, 2), 'utf8');

    // Schedule the task
    cron.schedule(cronTime, async () => {
      console.log(`Sending scheduled emails at ${scheduledDate.format()}`);

      try {
        // Remove and process the task
        let tasks = JSON.parse(fs.readFileSync(scheduledEmailsFilePath, 'utf8'));
        const taskIndex = tasks.findIndex(t => t.cronTime === cronTime);
        
        if (taskIndex === -1) {
          console.log('Task not found.');
          return;
        }

        const taskData = tasks.splice(taskIndex, 1)[0];
        fs.writeFileSync(scheduledEmailsFilePath, JSON.stringify(tasks, null, 2), 'utf8');
        
        const sheetData = await getSheetData(taskData.sheetId, taskData.sheetName);
        await sendEmails(sheetData, taskData.emailSubject, taskData.emailBody, taskData.attachment, taskData.ranges);

      } catch (error) {
        console.error('Error while sending scheduled emails:', error);
      }
    });

    res.json({ status: 'success', message: `Emails scheduled successfully for ${scheduledDate.format()}` });
  } catch (error) {
    console.error('Error in /schedule-emails:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while scheduling emails.' });
  }
});
// Function to get data from Google Sheets
async function getSheetData(sheetId, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: sheetName,
  });
  return response.data.values;
}

// Function to send emails with a delay
async function sendEmails(sheetData, emailSubject, emailBody, attachment, ranges) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const email = row[1]?.toString() || '';

    if (!name || !email) {
      console.log(`Skipping row ${i + 1} due to missing data.`);
      continue;
    }

    console.log(`Sending email ${i} of ${sheetData.length - 1} to ${name}...`);

    const formattedEmailBody = emailBody.replace('{{Name}}', name);

    const emailIndex = getEmailIndexForRange(i, ranges);
    if (emailIndex === -1) {
      console.log(`No email account defined for email ${i}. Skipping...`);
      continue;
    }

    await sendEmail(email, emailSubject, formattedEmailBody, attachment, emailIndex);
    await new Promise(resolve => setTimeout(resolve, 5000));
  }
}

function getEmailIndexForRange(emailNumber, ranges) {
  for (let i = 0; i < ranges.length; i++) {
    const { from, to } = ranges[i];
    if (emailNumber >= from && emailNumber <= to) {
      return i;
    }
  }
  return -1;
}

async function sendEmail(to, subject, htmlContent, attachment, emailIndex) {
  const transporter = getTransporter(emailIndex);

  const mailOptions = {
    from: transporter.options.auth.user,
    to,
    subject,
    html: htmlContent,
    attachments: [
      {
        filename: attachment.filename,
        content: attachment.content,
        encoding: 'base64',
        contentType: attachment.contentType
      }
    ]
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

// Function to get the transporter with credentials from the file
function getTransporter(emailIndex) {
  const emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
  const account = emailCredentials[emailIndex];

  if (!account) {
    throw new Error('Email account not found.');
  }

  return nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: account.user,
      pass: account.pass
    }
  });
}

app.listen(port, () => {
  console.log(`Server running on port no:${port}`);
});
