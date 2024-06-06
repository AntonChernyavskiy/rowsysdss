const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

// Роут для отображения страницы администратора
app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html'));
});

// Роут для обработки запросов на обновление кнопок
app.post('/admin/update-buttons', (req, res) => {
    // В этой части кода вы будете обрабатывать запрос на обновление кнопок.
    // Это может включать чтение данных из тела запроса и сохранение новой конфигурации кнопок.
    // Здесь предполагается, что вы уже реализовали соответствующую логику.
    // В случае успеха вы можете отправить ответ об успешном обновлении или перенаправить пользователя обратно на страницу администратора.

    // Примерный код для чтения данных из тела запроса и сохранения их в файл:
    // const buttonData = req.body.buttonData;
    // fs.writeFile('button-config.json', JSON.stringify(buttonData), (err) => {
    //     if (err) {
    //         console.error(err);
    //         res.status(500).send('Error updating button configuration');
    //     } else {
    //         res.redirect('/admin');
    //     }
    // });
});

// Роут для получения списка файлов
app.get('/files', (req, res) => {
    const filesDirectory = path.join(__dirname, 'public', 'files');
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
