// const { Pool } = require('pg');

// const pool = new Pool({
//     user: 'postgres',
//     host: 'localhost',
//     database: 'dummy',
//     password: 'postgres',
//     port: 5432,
// });

// async function queryDatabase(query, params) {
//     const client = await pool.connect();
//     try {
//         const res = await client.query(query, params);
//         return res.rows;
//     } catch (err) {
//         console.error('Error executing query', err.stack);
//     } finally {
//         client.release();
//     }
// }

// async function testConnection() {
//     const client = await pool.connect();
//     try {
//         const res = await client.query('SELECT NOW()');
//         console.log('Connected to the database:', res.rows[0]);
//     } catch (err) {
//         console.error('Error connecting to the database', err.stack);
//     } finally {
//         client.release();
//     }
// }
// testConnection();




// module.exports = {
//     queryDatabase,
// };



const mysql = require('mysql2/promise');

const connection = mysql.createPool({
    host: '110.239.90.35',
    user: 'RTS',
    password: 'RTS@0808',
    database: 'avr',
    port: 3308,
})

async function queryDatabase(query, params = []) {
    const conn = await connection.getConnection();
    try {

        const [rows] = params.length > 0
            ? await conn.execute(query, params)
            : await conn.query(query);
        return rows;
    } catch (err) {
        console.error('Error executing query', err.stack);
    } finally {
        conn.release(); // Close the connection after the query is executed
    }
}

module.exports = {
    queryDatabase,
};

// async function getData() {
//     try {
//         const rows = await queryDatabase('SELECT * FROM transaksi_his');
//         console.log(rows);
//     }
//     catch (err) {
//         console.error('Error fetching data', err.stack);
//     }
// }

// getData();