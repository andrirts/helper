const crypto = require('crypto');

function generateRandomAlphanumeric(length) {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    const charactersLength = characters.length;
    let result = '';

    // Generate a random string
    const randomBytes = crypto.randomBytes(length);
    for (let i = 0; i < length; i++) {
        const randomIndex = randomBytes[i] % charactersLength;
        result += characters[randomIndex];
    }

    return result;
}

// Generate a random alphanumeric string of length 32
module.exports = {
    generateRandomAlphanumeric
}