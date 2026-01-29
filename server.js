const express = require('express');
const path = require('path');
const app = express();
const port = 3000;

// Cung cấp tệp tĩnh từ thư mục 'public' (cho CSS và JS)
app.use(express.static(path.join(__dirname, 'public')));

// Cung cấp tệp tĩnh từ thư mục gốc (cho index.html và các file khác trong thư mục gốc)
app.use(express.static(path.join(__dirname)));

// Route trang chủ
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Khởi động server
app.listen(port, () => {
    console.log(`App listening at http://localhost:${port}`);
});