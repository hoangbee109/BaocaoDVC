const xlsx = require("xlsx");
const { Pool } = require("pg");

const pool = new Pool({
user:"postgres",
host:"localhost",
database:"baocao",
port:5432
});

async function run(){
const workbook = xlsx.readFile("tthc.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(sheet,{header:1});
let thu_tu = 1;
for(let i=1;i<rows.length;i++){
const stt = rows[i][0];
const ten = rows[i][1];
if(!stt || !ten) continue;
let cap_dong="tthc";
if(stt==="A" || stt==="B"){
cap_dong="nhom";
}
if(typeof stt==="string" && stt.match(/^[IVX]+$/)){
cap_dong="linhvuc";
}
await pool.query(
`INSERT INTO chitieu(stt,ten_chitieu,cap_dong,thu_tu)
VALUES($1,$2,$3,$4)`,
[stt,ten,cap_dong,thu_tu]
);
thu_tu++;
}
console.log("Import thành công");
process.exit();
}
run();