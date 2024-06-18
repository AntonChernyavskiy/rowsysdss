const express = require('express');
const path = require('path');
const fs = require('fs').promises; // Using promises for async fs operations

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

// Serve admin.html on /admin route
app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html'));
});

// Serve directory listing of /public/files
app.get('/files', (req, res) => {
    const filesDirectory = path.join(__dirname, 'public', 'files');
    fs.readdir(filesDirectory)
        .then(files => res.json(files))
        .catch(err => {
            console.error('Error reading files directory:', err);
            res.status(500).send('Unable to scan files directory');
        });
});

// Serve directory listing of /public/files, similar to /files
app.get('/files2', (req, res) => {
    const filesDirectory = path.join(__dirname, 'public', 'files');
    fs.readdir(filesDirectory)
        .then(files => res.json(files))
        .catch(err => {
            console.error('Error reading files directory:', err);
            res.status(500).send('Unable to scan files directory');
        });
});

// Serve directory listing of /public/files/events/:eventName
app.get('/files/:eventName', (req, res) => {
    const eventName = req.params.eventName;
    const eventDirectory = path.join(__dirname, 'public', 'files', eventName);
    fs.readdir(eventDirectory)
        .then(files => res.json(files))
        .catch(err => {
            console.error('Error reading event directory:', err);
            res.status(500).send('Unable to scan event directory');
        });
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
