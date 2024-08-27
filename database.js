const { Pool } = require('pg');

const pool = new Pool({
    user: 'rts_user',
    host: '10.121.7.9',
    database: 'ppob_core',
    password: '3V8c2GDPd47wZEk4wGzBrqr3',
    port: 5433,
});

async function queryDatabase(query, params) {
    const client = await pool.connect();
    try {
        const res = await client.query(query, params);
        return res.rows;
    } catch (err) {
        console.error('Error executing query', err.stack);
    } finally {
        client.release();
    }
}

module.exports = {
    queryDatabase,
};
