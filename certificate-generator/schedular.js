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

const emailsfilePath = path.join(__dirname, 'luneblazeemails.json');
// File path to store email credentials
const emailCredentialsFilePath = path.join(__dirname, 'luneblaze_emailCredentials.json');

// File path to store scheduled email tasks
const scheduledEmailsFilePath = path.join(__dirname, 'luneblaze_scheduledEmails.json');

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
app.post('/add-email-credentials', (req, res) => {
  const { user, pass, alias } = req.body;

  if (!user || !pass) {
    return res.status(400).json({ status: 'error', message: 'Email and password are required.' });
  }

  // Read existing credentials
  let emailCredentials = [];
  if (fs.existsSync(emailCredentialsFilePath)) {
    emailCredentials = JSON.parse(fs.readFileSync(emailCredentialsFilePath, 'utf8'));
  }

  // Add new credentials
  emailCredentials.push({ user, pass, alias });

  // Save credentials to file
  fs.writeFileSync(emailCredentialsFilePath, JSON.stringify(emailCredentials, null, 2), 'utf8');

  res.json({ status: 'success', message: 'Email credentials added successfully.' });
});

// Route to edit email credentials (including alias)
app.put('/edit-email-credentials/:index', (req, res) => {
  const { user, pass, alias } = req.body;
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
    emailCredentials[index] = { user, pass, alias };

    // Save the updated credentials
    fs.writeFileSync(emailCredentialsFilePath, JSON.stringify(emailCredentials, null, 2), 'utf8');

    res.json({ status: 'success', message: 'Email credentials updated successfully.' });
  } catch (error) {
    console.error('Error editing email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while editing credentials.' });
  }
});

// Route to delete email credentials by index
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

  // Validate required fields
  if (!sheetId || !sheetName || !emailSubject || !emailBody || !ranges || !scheduledDateTime) {
    return res.status(400).json({ status: 'error', message: 'One or more parameters are missing.' });
  }

  try {
    // Validate that scheduled time is in the future
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

    // Handle attachments - single, multiple, or null
    let emailAttachments = null;
    if (attachment) {
      if (Array.isArray(attachment)) {
        // If it's an array, assign the array directly
        emailAttachments = attachment.map(att => ({
          filename: att.filename,
          content: att.content,
          contentType: att.contentType
        }));
      } else {
        // If it's a single object, wrap it in an array
        emailAttachments = [{
          filename: attachment.filename,
          content: attachment.content,
          contentType: attachment.contentType
        }];
      }
    }

    // Add new task
    const task = {
      sheetId,
      sheetName,
      emailSubject,
      emailBody,
      attachment: emailAttachments, // Store processed attachments
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
async function getSheetNames(sheetId) {
  try {
    // Fetch the spreadsheet metadata to get sheet names
    const response = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
    });

    // Extract sheet names from the response
    const sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
    return {
      status: 'success',
      sheetNames
    };
  } catch (error) {
    if (error.code === 403 || error.code === 401) {
      // Handle permission or authentication error
      return {
        status: 'error',
        message: 'Access Denied: Unable to connect to the spreadsheet. Please check your permissions.'
      };
    } else {
      // Handle any other errors
      return {
        status: 'error',
        message: 'An error occurred while fetching the spreadsheet data.',
        error
      };
    }
  }
}

// Fetch data from a specific sheet
async function getSheetData(sheetId, sheetName) {

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: sheetName,
    });
      return response.data.values;
  }
    
app.post('/get-sheet-names', async (req, res) => {
  const { sheetId } = req.body;

  if (!sheetId) {
    return res.status(400).json({ status: 'error', message: 'Sheet ID is required.' });
  }

  const result = await getSheetNames(sheetId);
  if (result.status === 'success') {
    res.json({ status: 'success', sheetNames: result.sheetNames });
  } else {
    res.status(400).json({ status: 'error', message: result.message });
  }
});

// Function to send emails with a delay
async function sendEmails(sheetData, emailSubject, emailBody, attachment, ranges) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const email = row[3]?.toString() || '';

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

    // Check if attachment is provided before sending
    if (attachment) {
      await sendEmail(email, emailSubject, formattedEmailBody, attachment, emailIndex);
    } else {
      await sendEmail(email, emailSubject, formattedEmailBody, null, emailIndex); // Pass null or handle it in sendEmail
    }

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

  // Determine whether to use alias or user for the 'from' field
  const fromAddress = transporter.options.auth.alias || transporter.options.auth.user;

  // Normalize HTML content to remove excessive whitespace
  const normalizedHtmlContent = htmlContent
    .replace(/\s+/g, ' ')  // Replace multiple spaces/newlines with a single space
    .trim(); // Remove leading and trailing spaces

  // Set up mail options
  const mailOptions = {
    from: fromAddress, // Use alias if available, otherwise use user email
    to,
    subject,
    html: normalizedHtmlContent, // Use the normalized HTML content
  };

  // Handle attachments: single, multiple, or null
  if (attachment) {
    if (Array.isArray(attachment)) {
      // Multiple attachments
      mailOptions.attachments = attachment.map(att => ({
        filename: att.filename,
        content: att.content,
        encoding: 'base64',
        contentType: att.contentType,
      }));
    } else {
      // Single attachment
      mailOptions.attachments = [{
        filename: attachment.filename,
        content: attachment.content,
        encoding: 'base64',
        contentType: attachment.contentType,
      }];
    }
  }

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
      pass: account.pass,
      alias:account.alias
    }
  });
}

// Get emails
app.get('/emails', (req, res) => {
  fs.readFile(emailsfilePath, 'utf8', (err, data) => {
    if (err) {
      return res.status(500).json({ error: 'Failed to read file' });
    }
    res.json(JSON.parse(data));
  });
});

// Add email
app.post('/pemails', (req, res) => {
  const newEmail = req.body.email;
  
  fs.readFile(emailsfilePath, 'utf8', (err, data) => {
    if (err) {
      return res.status(500).json({ error: 'Failed to read file' });
    }
    
    let emails = JSON.parse(data);
    if (!emails.includes(newEmail)) {
      emails.push(newEmail);
      fs.writeFile(emailsfilePath, JSON.stringify(emails), (err) => {
        if (err) {
          return res.status(500).json({ error: 'Failed to write file' });
        }
        res.status(201).json({ message: 'Email added' });
      });
    } else {
      res.status(400).json({ error: 'Email already exists' });
    }
  });
});

// Delete email
app.delete('/demails', (req, res) => {
  const emailToDelete = req.body.email;
  
  fs.readFile(emailsfilePath, 'utf8', (err, data) => {
    if (err) {
      return res.status(500).json({ error: 'Failed to read file' });
    }
    
    let emails = JSON.parse(data);
    emails = emails.filter(email => email !== emailToDelete);
    
    fs.writeFile(emailsfilePath, JSON.stringify(emails), (err) => {
      if (err) {
        return res.status(500).json({ error: 'Failed to write file' });
      }
      res.json({ message: 'Email deleted' });
    });
  });
});

app.listen(port, () => {
  console.log(`Server running on port no:${port}`);
});
