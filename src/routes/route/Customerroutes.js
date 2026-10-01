const express = require('express');
const router = express.Router();
const { getCustomers } = require('../controllers/customerController');
const { getPartner } = require('../controllers/PartnerController');

// 0. Customer
router.get('/cust', getCustomers);
// 1. Partner
router.get('/partner', getPartner)

module.exports = router;