const express = require('express');
const router = express.Router();
const { verifyToken } = require('../middleware/auth');
const messageController = require('../controllers/messageController');

// 2. MESSAGES
// GET: butuh login (admin & developer melihat semua pesan, role lain hanya pesan miliknya)
// POST: butuh login (semua user yang sudah punya token boleh kirim pesan)
router.get('/messages', verifyToken, messageController.getMessages);
router.post('/messages', verifyToken, messageController.createMessage);

module.exports = router;