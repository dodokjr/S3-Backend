const { customers } = require('../Data/customers');

// GET /customers (terbuka, tanpa token)
exports.getCustomers = (req, res, next) => {
  try {
    res.json({ success: true, data: customers });
  } catch (error) {
    next(error);
  }
};