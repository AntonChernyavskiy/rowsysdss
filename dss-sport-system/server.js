const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html'));
});

app.get('/files', (req, res) => {
    const filesDirectory = path.join(__dirname, 'public', 'files');
    fs.readdir(filesDirectory, (err, files) => {
        if (err) {
            return res.status(500).send('Unable to scan files directory');
        }
        res.json(files);
    });
});

app.get('/files2', (req, res) => {
    const filesDirectory = path.join(__dirname, 'public', 'files', 'events');
    fs.readdir(filesDirectory, (err, files) => {
        if (err) {
            return res.status(500).send('Unable to scan files directory');
        }
        res.json(files);
    });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
