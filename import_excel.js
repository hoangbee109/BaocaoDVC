const xlsx = require("xlsx");
const { Pool } = require("pg");

/* kết nối database */

const pool = new Pool({
user:"postgres",
host:"localhost",
database:"baocao",
port:5432
});

/* đọc file excel */

const workbook = xlsx.readFile("danh_sach_tthc.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet,{header:1});

/* hàm kiểm tra số la mã */

function isRoman(text){
return /^(I|II|III|IV|V|VI|VII|VIII|IX|X)$/.test(text);
}

/* hàm kiểm tra số */

function isNumber(text){
return /^[0-9]+$/.test(text);
}

async function run(){
await pool.query("DELETE FROM chitieu");
let currentGroup = "";
let currentField = "";

for(let i=0;i<data.length;i++){
const row = data[i];
const stt = String(row[0] || "").trim();
const name = String(row[1] || "").trim();
if(!stt && !name) continue;

/* nhóm A,B */

if(stt==="A" || stt==="B"){
currentGroup = name;
await pool.query(
`INSERT INTO chitieu(stt,ten_chitieu,cap_dong)
VALUES($1,$2,'nhom')`,
[stt,name]
);
continue;
}

/* lĩnh vực */

if(isRoman(stt)){
currentField = name;
await pool.query(
`INSERT INTO chitieu(stt,ten_chitieu,cap_dong)
VALUES($1,$2,'linhvuc')`,
[stt,name]
);
continue;
}

/* thủ tục hành chính */

if(isNumber(stt)){
await pool.query(
`INSERT INTO chitieu(stt,ten_chitieu,cap_dong)
VALUES($1,$2,'tthc')`,
[stt,name]
);
}
}

console.log("Import hoàn thành");
process.exit();
}

run();