const express = require('express');
const app = express();
const { PORT } = require('./config/config');
const documentRoutes = require('./routes/documentRoutes');

app.use(express.json());

app.use('/generate-word', documentRoutes);

app.listen(PORT, () => {
    console.log(`API escuchando en http://localhost:${PORT}`);
});
