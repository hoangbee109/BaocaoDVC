const xlsx = require('xlsx');
const { Pool } = require('pg');

const pool = new Pool({
user: 'postgres',
host: 'localhost',
database: 'baocao',
port: 5432,
});

async function importXa(){

try{

/* đọc file excel */
const workbook = xlsx.readFile('xa.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

console.log("Đang import:", data.length, "xã");

/* 🔥 XOÁ DỮ LIỆU CŨ */
await pool.query("DELETE FROM xa");

/* reset id */
await pool.query("ALTER SEQUENCE xa_id_seq RESTART WITH 1");

/* insert lại */
for(const row of data){

await pool.query(
"INSERT INTO xa(id, ma_xa, ten_xa) VALUES($1,$2,$3)",
[row.id, row.ma_xa, row.ten_xa]
);

}

console.log("✅ Import xã thành công");

}catch(err){
console.error("❌ Lỗi:", err);
}

process.exit();

}

importXa();