require('dotenv').config();

module.exports = {
    PORT: process.env.PORT || 3000,
    SERVER_URL: process.env.SERVER_URL || 'http://localhost'
};
