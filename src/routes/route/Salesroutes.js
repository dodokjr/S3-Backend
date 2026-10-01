const express = require('express');
const router = express.Router();
const { verifyToken, checkModuleAccess } = require('../middleware/auth');
const salesController = require('../controllers/Salescontroller');

// 4. SALES (sinkron dengan stock)
// Akses GET/POST/PUT/DELETE: Sales, Developer, Admin
router.get('/sales', verifyToken, checkModuleAccess('sales'), salesController.getSales);
router.post('/sales', verifyToken, checkModuleAccess('sales'), salesController.createSales);
router.put('/sales', verifyToken, checkModuleAccess('sales'), salesController.updateSales);
router.delete('/sales', verifyToken, checkModuleAccess('sales'), salesController.deleteSales);

module.exports = router;