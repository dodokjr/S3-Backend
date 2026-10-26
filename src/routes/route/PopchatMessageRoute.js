const express = require('express');
const router = express.Router();

// Sesuaikan path relatif ini dengan struktur foldermu
const ctrl = require('../controllers/Popchatmessagecontroller');
const { verifyToken, allowCustomer, allowDeveloperAndAdmin, checkModuleAccess } = require('../middleware/auth');

// ==========================================
// POP CHAT ROUTES
// ==========================================
// Akses:
//   POST   /popchat/token  -> publik (customer minta token dengan nama & email)
//   POST   /popchat        -> customer saja (body: { pesan }, nama & email dari token)
//   GET    /popchat        -> customer (hanya chat miliknya) & staff (semua / ?key=)
//   POST   /popchat/customer-reply -> customer saja (membalas pesan CS/admin/developer)
//   POST   /popchat/reply  -> developer, admin & cs
//   DELETE /popchat        -> HANYA developer & admin

router.post('/popchat/token', ctrl.createCustomerToken);

router.post('/popchat', verifyToken, allowCustomer, ctrl.createChat);

router.get('/popchat', verifyToken, ctrl.getChats);

router.post('/popchat/customer-reply', verifyToken, allowCustomer, ctrl.customerReply);

router.post('/popchat/reply', verifyToken, checkModuleAccess('popchat'), ctrl.replyChat);

router.delete('/popchat', verifyToken, allowDeveloperAndAdmin, ctrl.deleteChat);

module.exports = router;