const express = require('express');
const router = express.Router();
const { verifyToken, checkModuleAccess } = require('../middleware/auth');
const financeController = require('../controllers/Financecontroller');

// 3. KEUANGAN (FINANCIAL)
// GET: terbuka (tanpa token)
// POST/PUT/DELETE: Finance, Developer, Admin
router.get('/finance', financeController.getFinance);
router.post('/finance', verifyToken, checkModuleAccess('finance'), financeController.createFinance);
router.put('/finance', verifyToken, checkModuleAccess('finance'), financeController.updateFinance);
router.delete('/finance', verifyToken, checkModuleAccess('finance'), financeController.deleteFinance);

module.exports = router;