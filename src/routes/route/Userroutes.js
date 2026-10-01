const express = require('express');
const router = express.Router();
const { verifyToken, allowDeveloperAndAdmin } = require('../middleware/auth');
const userController = require('../controllers/Usercontroller');

// 1. USERS
// GET: butuh login (sales/finance/gudang hanya melihat profil sendiri)
// POST/PUT/DELETE: Developer & Admin
router.get('/users', verifyToken, userController.getUsers);
router.post('/users', verifyToken, allowDeveloperAndAdmin, userController.createUser);
router.put('/users', verifyToken, allowDeveloperAndAdmin, userController.updateUser);
router.delete('/users', verifyToken, allowDeveloperAndAdmin, userController.deleteUser);

module.exports = router;