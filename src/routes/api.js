const express = require('express');
const router = express.Router();

// Gabungan semua route. Di-mount di app.js: app.use('/s3/api', apiRoutes);
router.use(require('./route/Customerroutes'));
router.use(require('./route/Userroutes'));
router.use(require('./route/Stockroutes'));
router.use(require('./route/Financeroutes'));
router.use(require('./route/Salesroutes'));

module.exports = router;