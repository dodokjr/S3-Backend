const { Partner } = require('../Data/Partner');

// GET /customers (terbuka, tanpa token)
exports.getPartner = (req, res, next) => {
  try {
    res.json({ success: true, data: Partner });
  } catch (error) {
    next(error);
  }
};