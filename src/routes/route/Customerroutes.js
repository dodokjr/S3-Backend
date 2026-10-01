const express = require('express');
const router = express.Router();
const { getCustomers } = require('../controllers/customerController');

// 0. Customer
router.get('/customers', getCustomers);

module.exports = router;