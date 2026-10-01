const express = require('express');
const router = express.Router();
const { verifyToken, checkModuleAccess } = require('../middleware/auth');
const stockController = require('../controllers/Stockcontroller');

// 2. REAL STOCK
// GET /stock: terbuka (tanpa token), mengirim SEMUA data; hidden = true jika PerPcs <= 0
// POST/PUT/DELETE: Gudang, Sales, Admin, Developer
router.get('/stock', stockController.getStock);
router.post('/stock', verifyToken, checkModuleAccess('gudang'), stockController.createStock);
router.put('/stock', verifyToken, checkModuleAccess('gudang'), stockController.updateStock);
router.delete('/stock', verifyToken, checkModuleAccess('gudang'), stockController.deleteStock);

module.exports = router;