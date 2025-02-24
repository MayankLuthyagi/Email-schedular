const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const dayjs = require('dayjs');
const {google} = require("googleapis");
const bodyParser = require("body-parser");
const express = require("express");
require('dotenv').config();

const scheduledEmailSchema = new mongoose.Schema({
    main: String,
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

const app = express();
mongoose.connect(process.env.MONGODB_URI)
    .then(() => console.log('MongoDB connected'))
    .catch(err => console.error('MongoDB connection error:', err));
const AutoIncrement = require('mongoose-sequence')(mongoose);
const ScheduledEmail = mongoose.model('ScheduledEmail', scheduledEmailSchema);

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

const scheduleTasks = async () => {
    try {
        // Fetch emails that should be sent now or earlier
        const now = dayjs().toISOString();
        const scheduledEmails = await ScheduledEmail.find({ scheduledDateTime: { $lte: now } });

        if (scheduledEmails.length === 0) {
            console.log('No pending emails to send.');
            await mongoose.connection.close(); // Close the connection before returning
            console.log('MongoDB connection closed.');
            return;
        }

        // Process each scheduled email in parallel
        await Promise.all(scheduledEmails.map(async (task) => {
            console.log(`Sending email scheduled at ${task.scheduledDateTime}`);

            try {
                // Send the email
                await sendEmails(
                    task.main, task.sheetId, task.sheetName,
                    task.emailId, task.pass, task.alias,
                    task.emailSubject, task.emailBody,
                    task.attachment, task.ranges
                );

                // Remove the task after execution
                await ScheduledEmail.deleteOne({ _id: task._id });
                console.log(`Email sent and task deleted for ${task.scheduledDateTime}`);
            } catch (error) {
                console.error(`Error sending scheduled email: ${error}`);
            }
        }));

        // Close MongoDB connection after processing all emails
        await mongoose.connection.close();
        console.log('MongoDB connection closed.');
    } catch (error) {
        console.error('Error in scheduling tasks:', error);
        await mongoose.connection.close(); // Ensure connection closes on error
        console.log('MongoDB connection closed due to an error.');
    }
};
async function sendEmails(main, sheetId, sheetName, emailId, pass, alias, emailSubject, emailBody, attachment, ranges) {
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

        const len = Math.min(ranges[1] + 1, rows.length - 1);
        const start = Math.max(ranges[0], 1);
        const array_email = [];
        console.log(`Schedule mailed in ranged ${start}-${len}`);
        for (let i = start; i <= len; i++) {
            const row = rows[i];
            const email = row[0]?.toString() || ''; // Convert to string, default to empty string if undefined
            if (!email) {
                console.log(`Skipping row ${i} in google due to missing data.`);
                continue;
            }
            array_email.push(email);
        }
        if (array_email.length === 0) {
            return console.error('Google Sheet is empty');
        }
        if (attachment) {
            await sendEmail(main, emailId, array_email, alias, pass, emailSubject, emailBody, attachment);
        } else {
            await sendEmail(main, emailId, array_email, alias, pass, emailSubject, emailBody, null);
        }
    } catch (error) {
        console.error('Error in sendEmails function:', error);
    }
}

async function sendEmail(main, emailId, bcc, alias, pass, subject, htmlContent, attachment) {
    try {
        const transporter = await getTransporter(main, alias, pass);
        const fromAddress = transporter.options.auth.alias || transporter.options.auth.user;

        const mailOptions = {
            from: fromAddress,
            to: emailId || undefined,
            bcc: Array.isArray(bcc) ? bcc.join(", ") : bcc,
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
        console.log(`Email sent successfully`);
    } catch (error) {
        console.error('Error in sendEmail:', error);
    }
}

async function getTransporter(email, alias, pass) {
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

scheduleTasks();
