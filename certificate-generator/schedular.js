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
const port = 3000;

// MongoDB connection
mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('MongoDB connected'))
  .catch(err => console.error('MongoDB connection error:', err));

// MongoDB schemas
const scheduledEmailSchema = new mongoose.Schema({
  sheetId: String,
  sheetName: String,
  emailId: String,
  pass: String,
  alias: String,
  emailSubject: String,
  emailBody: String,
  attachment: Array,
  ranges: Array,
  scheduledDateTime: String,
  cronTime: String,
});
const emailAdmin = new mongoose.Schema({
  authEmail: String,
});
const emailListSchema = new mongoose.Schema({
  email: String,
  sheetId: String,
  sheetName: String,
  pass: String,
  alias: String,
  min: Number,
  max: Number
});

// MongoDB models

const ScheduledEmail = mongoose.model('ScheduledEmail', scheduledEmailSchema);
const EmailList = mongoose.model('EmailList', emailListSchema);
const EmailAdmin = mongoose.model('EmailAdmin', emailAdmin);

// CORS configuration
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'email']
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
      scheduledDateTime: task.scheduledDateTime,
      alias: task.alias
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



// Route to get email addresses
app.get('/get-emails', async (req, res) => {
  try {
    const emailCredentials = req.headers['email'];
    res.json({ emails: emailCredentials });
  } catch (error) {
    console.error('Error fetching email accounts:', error);
    res.status(500).json({ message: 'Failed to fetch email accounts' });
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

app.post('/schedule-emails', async (req, res) => {
  const { sheetId, sheetName, email, pass, alias, emailSubject, emailBody, attachment, ranges, scheduledDateTime } = req.body;

  // Validate required fields
  if (!sheetId || !sheetName || !pass || !email || !emailSubject || !emailBody || !ranges || !scheduledDateTime) {
    return res.status(400).json({ status: 'error', message: `One or more parameters are missing.` });
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
      emailId: email,
      pass,
      alias,
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
      console.log(`Scheduled job triggered at ${scheduledDate.format()}`);

      const taskData = await ScheduledEmail.findOne({ cronTime });
      if (!taskData) {
        console.log('Task not found.');
        return;
      }

      await sendEmails(taskData.sheetId, taskData.sheetName, taskData.emailId, taskData.pass, taskData.alias, taskData.emailSubject, taskData.emailBody, taskData.attachment, taskData.ranges);
      await ScheduledEmail.deleteOne({ cronTime });
    });

    res.json({ status: 'success', message: `Emails scheduled successfully for ${scheduledDate.format()}` });
  } catch (error) {
    console.error('Error in /schedule-emails:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while scheduling emails.' });
  }
});

async function sendEmails(sheetId, sheetName, emailId, pass, alias, emailSubject, emailBody, attachment, ranges) {
  try {
    if (!sheetId || !sheetName) {
      throw new Error('No data found in the Google Sheet.');
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: sheetName,
    });

    const rows = response.data.values;
    if (!rows || rows.length < 2) {
      console.log("No data found or only headers exist.");
      return;  // Exit early if no data is found
    }

    const len = Math.min(ranges[1], rows.length - 1);
    const start = Math.max(ranges[0], 1);

    for (let i = start; i <= len; i++) {
      const row = rows[i];
      const email = row[0]?.toString() || ''; // Convert to string, default to empty string if undefined

      if (!email) {
        console.log(`Skipping row ${i} due to missing data.`);
        continue;
      }

      console.log(`Sending email ${i} of ${len}`);

      // Check if attachment is provided before sending
      if (attachment) {
        await sendEmail(email, emailId, alias, pass, emailSubject, emailBody, attachment);
      } else {
        await sendEmail(email, emailId, alias, pass, emailSubject, emailBody, null);
      }

      await new Promise(resolve => setTimeout(resolve, 5000));  // Wait for 5 seconds before sending the next email
    }
  } catch (error) {
    console.error('Error in sendEmails function:', error);
  }
}

async function sendEmail(to, from, alias, pass, subject, htmlContent, attachment) {
  try {
    const transporter = await getTransporter(from, pass, alias);
    const fromAddress = transporter.options.auth.alias || transporter.options.auth.user;

    const mailOptions = {
      from: fromAddress,
      to,
      subject,
      html: htmlContent,
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
    console.error('Error in sendEmail:', error);
  }
}

async function getTransporter(email, pass, alias) {
  try {

    if (!email || !pass) {
      throw new Error(`Email account not found or missing credentials for: ${email}`);
    }

    return nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: email,
        pass: pass,
        alias: alias
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
app.get('/alias', async (req, res) => {
  try {
    const alias = await EmailList.find();
    const aliasAddresses = alias.map(emailEntry => emailEntry.alias);
    res.json(aliasAddresses);
  } catch (error) {
    console.error('Error fetching alias:', error);
    res.status(500).json({ error: 'Failed to fetch alias' });
  }
});

app.get('/authemails', async (req, res) => {
  try {
    const emails = await EmailAdmin.find();
    const emailAddresses = emails.map(emailEntry => emailEntry.authEmail);
    res.json(emailAddresses);
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ error: 'Failed to fetch emails' });
  }
});

app.get('/sheets-detail', async (req, res) => {
  try {
    const data = await EmailList.find(); // Fetch all email entries from the database
    res.json(data); // Send the complete data as JSON response
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ error: 'Failed to fetch emails' });
  }
});


app.get('/credential-details', async (req, res) => {
  try {
    const userEmail = req.headers['email'];  // Retrieve email from the header

    const sheet = await EmailList.findOne({ email: userEmail });
    if (sheet) {
      return res.json({
        email: sheet.email,
        pass: sheet.pass,
        alias: sheet.alias
      });
    } else {
      return res.status(404).json({ message: "No user found with this email." });
    }
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ error: 'Failed to fetch emails' });
  }
});

// Route to add a new email to the email list collection
app.post('/pemails', async (req, res) => {
  const { email, sheetId, sheetName, pass, alias, min, max } = req.body;

  if (!email || !sheetId || !sheetName || !min || !max || !pass || !alias) {
    return res.status(400).json({ error: 'Email is required' });
  }

  try {
    // Check if the email already exists in the database
    const existingEmail = await EmailList.findOne({ email: email });
    if (existingEmail) {
      return res.status(400).json({ error: 'Email already exists' });
    }

    // Add the new email
    const emailEntry = new EmailList({ email, sheetId, sheetName, pass, alias, min, max });

    await emailEntry.save();

    res.status(201).json({ message: 'Email added' });
  } catch (error) {
    console.error('Error adding email:', error);
    res.status(500).json({ error: 'Failed to add email' });
  }
});
app.post('/pauthemails', async (req, res) => {
  const { authEmail } = req.body;

  if (!authEmail) {
    return res.status(400).json({ error: 'Email is required' });
  }

  try {
    // Check if the email already exists in the database
    const existingEmail = await EmailAdmin.findOne({ authEmail });
    if (existingEmail) {
      return res.status(400).json({ error: 'Email already exists' });
    }

    // Add the new email
    const emailEntry = new EmailAdmin({ authEmail });

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
app.delete('/dauthemails', async (req, res) => {
  const emailToDelete = req.body.authEmail;

  if (!emailToDelete) {
    return res.status(400).json({ error: 'Email is required' });
  }

  try {
    // Delete the email
    const result = await EmailAdmin.deleteOne({ authEmail: emailToDelete });

    if (result.deletedCount === 0) {
      return res.status(404).json({ error: 'Email not found' });
    }

    res.json({ message: 'Email deleted' });
  } catch (error) {
    console.error('Error deleting email:', error);
    res.status(500).json({ error: 'Failed to delete email' });
  }
});
app.put('/update-email', async (req, res) => {
  const { email, sheetId, sheetName, min, max, pass, alias } = req.body;

  if (!email || !sheetId || !sheetName || !min || !max || !pass || !alias) {
    return res.status(400).json({ status: 'error', message: 'All fields are required' });
  }

  try {
    // Update the email details in the EmailList collection
    const updatedEmail = await EmailList.findOneAndUpdate(
      { email }, // Search for the document with this email
      {
        sheetId,
        sheetName,
        min,
        max,
        pass,
        alias
      }, // Fields to update
      { new: true } // Return the updated document
    );

    if (!updatedEmail) {
      return res.status(404).json({ status: 'error', message: 'Email not found' });
    }

    res.status(200).json({ status: 'success', message: 'Email updated successfully', updatedEmail });
  } catch (error) {
    console.error(error);
    res.status(500).json({ status: 'error', message: 'Failed to update email' });
  }
});



cron.schedule('*/5 * * * *', () => {
  console.log('server is running new');
});

app.listen(port, () => {
  console.log(`Server running on port no:${port}`);
});

