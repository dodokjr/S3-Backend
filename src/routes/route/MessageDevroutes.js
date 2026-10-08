const express = require('express');
const router = express.Router();
const { verifyToken } = require('../middleware/auth');
const messageController = require('../controllers/MessageDevcontroller');

// 2. MESSAGES
// GET: butuh login (admin & developer melihat semua pesan, role lain hanya pesan miliknya)
//      Query opsional: ?tag=developer&status=pending
// POST: butuh login (semua user yang sudah punya token boleh kirim pesan, status awal: unread)
// PATCH: butuh login + role admin/developer (ubah status: unread | read | pending | success)
router.get('/messages', verifyToken, messageController.getMessages);
router.post('/messages', verifyToken, messageController.createMessage);
router.patch('/messages/:id/status', verifyToken, messageController.updateMessageStatus);

module.exports = router;