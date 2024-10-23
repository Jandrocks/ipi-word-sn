const express = require('express');
const router = express.Router();
const { generateWordDocument } = require('../controllers/documentController');

router.post('/', generateWordDocument);

// Nueva ruta GET para verificar el estado de la API
router.get('/status', (req, res) => {
    res.status(200).json({ message: 'API operativa' });
});

module.exports = router;
