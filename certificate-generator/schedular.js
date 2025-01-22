const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const bodyParser = require('body-parser');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const cron = require('node-cron');
const dayjs = require('dayjs');
require('dotenv').config();

const app = express();
const port = 3001;

// MongoDB connection
mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('MongoDB connected'))
  .catch(err => console.error('MongoDB connection error:', err));

// MongoDB schemas
const emailCredentialSchema = new mongoose.Schema({
  user: String,
  pass: String,
  alias: String,
});

const scheduledEmailSchema = new mongoose.Schema({
  sheetId: String,
  sheetName: String,
  emailSubject: String,
  emailBody: String,
  attachment: Array,
  ranges: Array,
  scheduledDateTime: String,
  cronTime: String,
});

const emailListSchema = new mongoose.Schema({
  email: String,
});

// MongoDB models
const EmailCredential = mongoose.model('EmailCredential', emailCredentialSchema);
const ScheduledEmail = mongoose.model('ScheduledEmail', scheduledEmailSchema);
const EmailList = mongoose.model('EmailList', emailListSchema);

// CORS configuration
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
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
app.get('/scheduled-tasks', async (req, res) => {
  try {
    const tasks = await ScheduledEmail.find();
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
app.delete('/delete-scheduled-task/:index', async (req, res) => {
  const { index } = req.params;

  try {
    const tasks = await ScheduledEmail.find();
    if (index < 0 || index >= tasks.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Delete the task
    await ScheduledEmail.deleteOne({ _id: tasks[index]._id });

    res.json({ status: 'success', message: 'Scheduled task deleted successfully.' });
  } catch (error) {
    console.error('Error deleting scheduled task:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while deleting the scheduled task.' });
  }
});

// Route to add email credentials
app.post('/add-email-credentials', async (req, res) => {
  const { user, pass, alias } = req.body;

  if (!user || !pass) {
    return res.status(400).json({ status: 'error', message: 'Email and password are required.' });
  }

  try {
    const emailCredential = new EmailCredential({ user, pass, alias });
    await emailCredential.save();
    res.json({ status: 'success', message: 'Email credentials added successfully.' });
  } catch (error) {
    console.error('Error adding email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while adding credentials.' });
  }
});

// Route to edit email credentials by index
app.put('/edit-email-credentials/:index', async (req, res) => {
  const { user, pass, alias } = req.body;
  const { index } = req.params;

  if (!user || !pass) {
    return res.status(400).json({ status: 'error', message: 'Email and password are required.' });
  }

  try {
    const emailCredentials = await EmailCredential.find();
    if (index < 0 || index >= emailCredentials.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Update credentials
    await EmailCredential.updateOne({ _id: emailCredentials[index]._id }, { user, pass, alias });

    res.json({ status: 'success', message: 'Email credentials updated successfully.' });
  } catch (error) {
    console.error('Error editing email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while editing credentials.' });
  }
});

// Route to delete email credentials by index
app.delete('/delete-email-credentials/:index', async (req, res) => {
  const { index } = req.params;

  try {
    const emailCredentials = await EmailCredential.find();
    if (index < 0 || index >= emailCredentials.length) {
      return res.status(400).json({ status: 'error', message: 'Invalid index.' });
    }

    // Delete the credentials
    await EmailCredential.deleteOne({ _id: emailCredentials[index]._id });

    res.json({ status: 'success', message: 'Email credentials deleted successfully.' });
  } catch (error) {
    console.error('Error deleting email credentials:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while deleting credentials.' });
  }
});

// Route to get email addresses
app.get('/get-emails', async (req, res) => {
  try {
    const emailCredentials = await EmailCredential.find();
    const emailAddresses = emailCredentials.map(account => account.user);
    res.json({ emails: emailAddresses });
  } catch (error) {
    console.error('Error fetching email accounts:', error);
    res.status(500).json({ message: 'Failed to fetch email accounts' });
  }
});

app.post('/schedule-emails', async (req, res) => {
  const { sheetId, sheetName, emailSubject, emailBody, attachment, ranges, scheduledDateTime } = req.body;

  // Validate required fields
  if (!sheetId || !sheetName || !emailSubject || !emailBody || !ranges || !scheduledDateTime) {
    return res.status(400).json({ status: 'error', message: 'One or more parameters are missing.' });
  }

  try {
    // Validate that the scheduled time is in the future
    const scheduledDate = dayjs(scheduledDateTime);
    const now = dayjs();
    if (scheduledDate.isBefore(now)) {
      return res.status(400).json({ status: 'error', message: 'Scheduled time must be in the future.' });
    }

    // Convert scheduledDate to cron format
    const cronTime = `${scheduledDate.second()} ${scheduledDate.minute()} ${scheduledDate.hour()} ${scheduledDate.date()} ${scheduledDate.month() + 1} *`;

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
    const newScheduledEmail = new ScheduledEmail({
      sheetId,
      sheetName,
      emailSubject,
      emailBody,
      attachment: emailAttachments, // Store processed attachments
      ranges,
      scheduledDateTime,
      cronTime
    });

    // Save the scheduled email to the database
    await newScheduledEmail.save();

    // Schedule the cron job to send emails at the specified time
    cron.schedule(cronTime, async () => {
      console.log(`Sending scheduled emails at ${scheduledDate.format()}`);

      const taskData = await ScheduledEmail.findOne({ cronTime });
      if (!taskData) {
        console.log('Task not found.');
        return;
      }

      const sheetData = await getSheetData(taskData.sheetId, taskData.sheetName);
      await sendEmails(sheetData, taskData.emailSubject, taskData.emailBody, taskData.attachment, taskData.ranges);

      // Remove the task after sending emails
      await ScheduledEmail.deleteOne({ cronTime });
    });

    res.json({ status: 'success', message: `Emails scheduled successfully for ${scheduledDate.format()}` });
  } catch (error) {
    console.error('Error in /schedule-emails:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while scheduling emails.' });
  }
});

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

async function sendEmails(sheetData, emailSubject, emailBody, attachment, ranges) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const email = row[0]?.toString() || '';

    if (!email) {
      console.log(`Skipping row ${i + 1} due to missing data.`);
      continue;
    }

    console.log(`Sending email ${i} of ${sheetData.length - 1}`);


    const emailIndex = getEmailIndexForRange(i, ranges);
    if (emailIndex === -1) {
      console.log(`No email account defined for email ${i}. Skipping...`);
      continue;
    }

    // Check if attachment is provided before sending
    if (attachment) {
      await sendEmail(email, emailSubject, emailBody, attachment, emailIndex);
    } else {
      await sendEmail(email, emailSubject, emailBody, null, emailIndex); // Pass null or handle it in sendEmail
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
  try {
    const transporter = await getTransporter(emailIndex);

    const fromAddress = transporter.options.auth.alias || transporter.options.auth.user;

    const normalizedHtmlContent = htmlContent
      .replace(/\n+/g, '<br>')
      .replace(/\s+/g, ' ')
      .trim();

    const mailOptions = {
      from: fromAddress,
      to,
      subject,
      html: `
        <div style="line-height: 1.5; max-width: 600px; margin: auto; font-size: 16px;">
          ${normalizedHtmlContent.replace(/<ul>/g, '<ul style="line-height: 1.5;">')
                                 .replace(/<ol>/g, '<ol style="line-height: 1.5;">')}
        </div>
        <style>
          @media only screen and (max-width: 600px) {
            div {
              line-height: 1.8; 
              font-size: 14px;
              word-wrap: break-word; /* Ensures long words or URLs break properly */
            }
          }
          p {
            margin: 1em 0; /* Ensure paragraphs have consistent spacing */
          }
        </style>`,
    };

    if (attachment) {
      mailOptions.attachments = Array.isArray(attachment)
        ? attachment.map(att => ({
            filename: att.filename,
            content: att.content,
            encoding: 'base64',
            contentType: att.contentType,
          }))
        : [{
            filename: attachment.filename,
            content: attachment.content,
            encoding: 'base64',
            contentType: attachment.contentType,
          }];
    }

    await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

async function getTransporter(emailIndex) {
  try {
    // Query the MongoDB collection to get email credentials by index
    const emailCredentials = await EmailCredential.find().skip(emailIndex).limit(1);
    const account = emailCredentials[0];  // Retrieve the first account

    if (!account) {
      throw new Error('Email account not found.');
    }

    return nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: account.user,
        pass: account.pass,
        alias: account.alias // Assuming alias is a field in the EmailCredentials schema
      }
    });
  } catch (error) {
    console.error('Error getting transporter:', error);
    throw error;
  }
}

// Route to get all emails from the email list collection
app.get('/emails', async (req, res) => {
  try {
    const emails = await EmailList.find();
    const emailAddresses = emails.map(emailEntry => emailEntry.email);
    res.json(emailAddresses);
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ error: 'Failed to fetch emails' });
  }
});

// Route to add a new email to the email list collection
app.post('/pemails', async (req, res) => {
  const newEmail = req.body.email;

  if (!newEmail) {
    return res.status(400).json({ error: 'Email is required' });
  }

  try {
    // Check if the email already exists in the database
    const existingEmail = await EmailList.findOne({ email: newEmail });
    if (existingEmail) {
      return res.status(400).json({ error: 'Email already exists' });
    }

    // Add the new email
    const emailEntry = new EmailList({ email: newEmail });
    await emailEntry.save();

    res.status(201).json({ message: 'Email added' });
  } catch (error) {
    console.error('Error adding email:', error);
    res.status(500).json({ error: 'Failed to add email' });
  }
});

// Route to delete an email from the email list collection
app.delete('/demails', async (req, res) => {
  const emailToDelete = req.body.email;

  if (!emailToDelete) {
    return res.status(400).json({ error: 'Email is required' });
  }

  try {
    // Delete the email
    const result = await EmailList.deleteOne({ email: emailToDelete });
    
    if (result.deletedCount === 0) {
      return res.status(404).json({ error: 'Email not found' });
    }

    res.json({ message: 'Email deleted' });
  } catch (error) {
    console.error('Error deleting email:', error);
    res.status(500).json({ error: 'Failed to delete email' });
  }
});

cron.schedule('*/5 * * * *', () => {
  console.log('server is running');
});

app.listen(port, () => {
  console.log(`Server running on port no:${port}`);
});

