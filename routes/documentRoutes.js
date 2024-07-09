const express = require('express');
const router = express.Router();
const { generateWordDocument } = require('../controllers/documentController');

router.post('/', generateWordDocument);

module.exports = router;
