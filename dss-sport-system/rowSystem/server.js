const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Обслуживание статических файлов из папок public и files
app.use(express.static(path.join(__dirname, 'public')));
app.use('/files', express.static(path.join(__dirname, 'files')));

// Маршрут для получения списка файлов
app.get('/files-list', (req, res) => {
    const filesDirectory = path.join(__dirname, 'files');
    fs.readdir(filesDirectory, (err, files) => {
        if (err) {
            return res.status(500).send('Unable to scan files directory');
        }
        res.json(files);
    });
});

// Запуск сервера
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
