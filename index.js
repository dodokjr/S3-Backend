require("dotenv").config();
const express = require('express');
const cors = require('cors');
const apiRoutes = require('./src/routes/api');
const authRoutes = require('./src/routes/auth');

const app = express();
app.use(express.json());
app.use(cors());

const PORT = process.env.PORT || 3001;

// Mendaftarkan Routes
app.use('/s3/api', apiRoutes);       // Mengakses /api/data dan /api/stock
app.use('/s3/api/auth', authRoutes); // Mengakses /api/auth/signin dan /api/auth/logout

app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});