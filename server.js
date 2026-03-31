function checkRole(role){
return (req,res,next)=>{
const roleClient = req.headers.role;
if(roleClient !== role){
return res.status(403).send("Không có quyền");
}
next();
};
}
const express = require("express");
const { Pool } = require("pg");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const app = express();

app.use(bodyParser.json());
app.use(express.static("public"));

/* Kết nối database */

const { Pool } = require('pg');

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false
  }
});

app.get("/ky-baocao", async(req,res)=>{
const result = await pool.query(
`
SELECT * FROM ky_baocao
WHERE trang_thai='mo'
ORDER BY id DESC
LIMIT 1
`
);
res.json(result.rows[0]);
});

app.get("/dashboard-xa", async(req,res)=>{
const thang = req.query.thang;
try{
const result = await pool.query(
`
SELECT x.id,x.ten_xa,
CASE
WHEN COUNT(d.id)>0
THEN 'Đã nhập'
ELSE 'Chưa nhập'
END trang_thai
FROM xa x
LEFT JOIN dulieu_baocao d
ON x.id=d.xa_id AND d.thang=$1
GROUP BY x.id,x.ten_xa
ORDER BY x.ten_xa
`,
[thang]
);
res.json(result.rows);
}catch(err){
console.error(err);
res.status(500).send("Lỗi server");
}
});

/*Xem log*/
app.get("/log", checkRole("admin"), async(req,res)=>{
const result = await pool.query(`
SELECT l.*,x.ten_xa
FROM log_hoatdong l
JOIN xa x ON l.xa_id=x.id
ORDER BY l.thoi_gian DESC
LIMIT 100
`);
res.json(result.rows);
});

/* reset mật khẩu*/
app.post("/reset-mat-khau", async(req,res)=>{

try{

const {xa_id} = req.body;

await pool.query(
"UPDATE users SET password='123456' WHERE xa_id=$1",
[xa_id]
);

res.json({success:true});

}catch(err){
console.error(err);
res.status(500).send("Lỗi server");
}

});

app.get("/chitieu", async (req,res)=>{
try{
const result = await pool.query(
"SELECT id,stt,ten_chitieu,cap_dong FROM chitieu WHERE stt <> 'TT' ORDER BY thu_tu"
);
res.json(result.rows);
}catch(err){ 
console.error(err);
res.status(500).send("Lỗi server");
}

});

app.get("/dulieu-xa", async (req,res)=>{

const xa_id = req.query.xa_id;
const thang = req.query.thang;

try{

const result = await pool.query(
`SELECT * FROM dulieu_baocao
WHERE xa_id=$1 AND thang=$2`,
[xa_id,thang]
);

res.json(result.rows);

}catch(err){

console.error(err);
res.status(500).send("Lỗi server");

}

});

app.post("/luu", checkRole("xa"), async (req,res)=>{

try{

/* khóa báo cáo sau ngày 11 */

const today = new Date();
const day = today.getDate();

if(day >= 32){
return res.status(403).send("Đã khóa báo cáo tháng này");
}

const data = req.body;
if(!data || data.length===0){
return res.send("OK");
}

/* xoá dữ liệu cũ */

await pool.query(
"DELETE FROM dulieu_baocao WHERE xa_id=$1 AND thang=$2",
[data[0].xa_id,data[0].thang]
);

/* chuẩn bị insert */

const values=[];

data.forEach(row=>{

values.push(
row.thang,
row.xa_id,
row.chitieu_id,

row.tiepnhan_tructuyen || 0,
row.tiepnhan_tructiep || 0,

row.giaiquyet_tructuyen || 0,
row.giaiquyet_tructiep || 0,

row.dangxuly_tructuyen || 0,
row.dangxuly_tructiep || 0,

row.thanhtoan_tructuyen || 0,

row.sohoa || 0,
row.lam_sach || 0,
row.chua_lam_sach || 0,

row.ho_so_qua_han || 0,
row.tra_kq_tructuyen || 0
);

});

/* ===== TỔNG TOÀN TỈNH ===== */

let sum_tn_tt=0;
let sum_tn_tp=0;

let sum_gq_tt=0;
let sum_gq_tp=0;

let sum_dx_tt=0;
let sum_dx_tp=0;

let sum_quahan=0;
let sum_trakq=0;
let sum_thanhtoan=0;

let sum_sohoa=0;
let sum_lamsach=0;
let sum_chualamsach=0;

/* duyệt lại data */
data.forEach(r=>{

sum_tn_tt += Number(r.tn_tt||0);
sum_tn_tp += Number(r.tn_tp||0);

sum_gq_tt += Number(r.gq_tt||0);
sum_gq_tp += Number(r.gq_tp||0);

sum_dx_tt += Number(r.dx_tt||0);
sum_dx_tp += Number(r.dx_tp||0);

sum_quahan += Number(r.qua_han||0);
sum_trakq += Number(r.tra_kq||0);

sum_thanhtoan += Number(r.thanhtoan||0);

sum_sohoa += Number(r.sohoa||0);
sum_lamsach += Number(r.lam_sach||0);
sum_chualamsach += Number(r.chualamsach||0);

});

/* tính tổng */
const sum_tn = sum_tn_tt + sum_tn_tp;
const sum_gq = sum_gq_tt + sum_gq_tp;
const sum_sohoa_total = sum_sohoa + sum_lamsach + sum_chualamsach;
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet();
/* thêm dòng tổng */
const totalRow = sheet.addRow([
'',
'TỔNG TOÀN TỈNH',

sum_tn_tt,
sum_tn_tp,
sum_tn,

sum_gq_tt,
sum_gq_tp,
sum_gq,

sum_dx_tt,
sum_dx_tp,

sum_quahan,
sum_trakq,
sum_thanhtoan,

sum_sohoa,
sum_lamsach,
sum_chualamsach,
sum_sohoa_total
]);

/* style dòng tổng */
totalRow.eachCell((cell, colNumber)=>{

cell.font = { bold:true };

cell.border = {
top:{style:'medium'},
left:{style:'thin'},
bottom:{style:'medium'},
right:{style:'thin'}
};

if(colNumber >= 3){
cell.numFmt = '#,##0';
cell.alignment = { horizontal:'center' };
}else{
cell.alignment = { horizontal:'left' };
}

/* màu xanh nhẹ */
cell.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'D4EDDA'}
};

});

/* tạo query */

const query = `
INSERT INTO dulieu_baocao
(thang,xa_id,chitieu_id,
tiepnhan_tructuyen,tiepnhan_tructiep,
giaiquyet_tructuyen,giaiquyet_tructiep,
dangxuly_tructuyen,dangxuly_tructiep,
thanhtoan_tructuyen,
sohoa,lam_sach,chua_lam_sach,trang_thai,ho_so_qua_han,
tra_kq_tructuyen)
VALUES
${data.map((_,i)=>{

const p=i*15;

return `($${p+1},$${p+2},$${p+3},
$${p+4},$${p+5},
$${p+6},$${p+7},
$${p+8},$${p+9},
$${p+10},
$${p+11},$${p+12},$${p+13},'draft',$${p+14}, $${p+15})`;
}).join(",")}
`;
await pool.query(query,values);
await pool.query(
"INSERT INTO log_hoatdong(xa_id,hanh_dong) VALUES($1,$2)",
[data[0].xa_id,"Lưu tạm báo cáo"]
);
res.send("OK");
}catch(err){
console.error(err);
res.status(500).send("Lỗi lưu dữ liệu");
}

});

app.get("/tonghop", async (req,res)=>{

try{

const thang = req.query.thang;

const result = await pool.query(`
SELECT 
c.id,
c.stt,
c.ten_chitieu,

COALESCE(SUM(d.tiepnhan_tructuyen),0) as tn_tt,
COALESCE(SUM(d.tiepnhan_tructiep),0) as tn_tp,

COALESCE(SUM(d.giaiquyet_tructuyen),0) as gq_tt,
COALESCE(SUM(d.giaiquyet_tructiep),0) as gq_tp,

COALESCE(SUM(d.dangxuly_tructuyen),0) as dx_tt,
COALESCE(SUM(d.dangxuly_tructiep),0) as dx_tp,

COALESCE(SUM(d.ho_so_qua_han),0) as qua_han,
COALESCE(SUM(d.tra_kq_tructuyen),0) as tra_kq,

COALESCE(SUM(d.thanhtoan_tructuyen),0) as thanhtoan,

COALESCE(SUM(d.sohoa),0) as sohoa,
COALESCE(SUM(d.lam_sach),0) as lam_sach,
COALESCE(SUM(d.chua_lam_sach),0) as chualamsach

FROM chitieu c
LEFT JOIN dulieu_baocao d 
ON c.id = d.chitieu_id AND d.thang = $1

GROUP BY c.id, c.stt, c.ten_chitieu
ORDER BY c.id
`,[thang]);

res.json(result.rows);

}catch(err){

console.error("Lỗi tonghop:", err);

/* 🔥 QUAN TRỌNG: trả JSON */
res.json([]);

}

});

app.get("/export-excel", async (req,res)=>{

try{

const thang = req.query.thang;
/* EXCEL */
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('TongHop');
/* QUERY */
const result = await pool.query(`
SELECT 
c.stt,
c.ten_chitieu,

COALESCE(SUM(d.tiepnhan_tructuyen),0) as tn_tt,
COALESCE(SUM(d.tiepnhan_tructiep),0) as tn_tp,

COALESCE(SUM(d.giaiquyet_tructuyen),0) as gq_tt,
COALESCE(SUM(d.giaiquyet_tructiep),0) as gq_tp,

COALESCE(SUM(d.dangxuly_tructuyen),0) as dx_tt,
COALESCE(SUM(d.dangxuly_tructiep),0) as dx_tp,

COALESCE(SUM(d.ho_so_qua_han),0) as qua_han,
COALESCE(SUM(d.tra_kq_tructuyen),0) as tra_kq,

COALESCE(SUM(d.thanhtoan_tructuyen),0) as thanhtoan,

COALESCE(SUM(d.sohoa),0) as sohoa,
COALESCE(SUM(d.lam_sach),0) as lam_sach,
COALESCE(SUM(d.chua_lam_sach),0) as chualamsach

FROM chitieu c
LEFT JOIN dulieu_baocao d 
ON c.id = d.chitieu_id AND d.thang = $1

GROUP BY c.stt, c.ten_chitieu, c.id
ORDER BY c.id
`,[thang]);

const data = result.rows;

/* HEADER đơn giản trước (test OK rồi nâng cấp sau) */
/* ===== HEADER 2 DÒNG ===== */

sheet.mergeCells('A1:A2');
sheet.getCell('A1').value = 'STT';

sheet.mergeCells('B1:B2');
sheet.getCell('B1').value = 'Dịch vụ công';

sheet.mergeCells('C1:E1');
sheet.getCell('C1').value = 'HỒ SƠ TIẾP NHẬN';

sheet.mergeCells('F1:H1');
sheet.getCell('F1').value = 'ĐÃ GIẢI QUYẾT';

sheet.mergeCells('I1:J1');
sheet.getCell('I1').value = 'ĐANG GIẢI QUYẾT';

sheet.mergeCells('K1:K2');
sheet.getCell('K1').value = 'QUÁ HẠN';

sheet.mergeCells('L1:L2');
sheet.getCell('L1').value = 'TRẢ KQ';

sheet.mergeCells('M1:M2');
sheet.getCell('M1').value = 'THANH TOÁN';

sheet.mergeCells('N1:Q1');
sheet.getCell('N1').value = 'SỐ HOÁ';

/* dòng 2 */
sheet.getRow(2).values = [
'', '',
'TT','TP','Tổng',
'TT','TP','Tổng',
'TT','TP',
'', '', '',
'Đã SH','Làm sạch','Chưa sạch','Tổng'
];

/* ===== STYLE HEADER ===== */

[1,2].forEach(rowIndex=>{
sheet.getRow(rowIndex).eachCell(cell=>{
cell.font = { bold: true };
cell.alignment = { vertical:'middle', horizontal:'center' };

cell.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'D9EEF3'}
};

cell.border = {
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};

});
});

/* DATA */
data.forEach(r=>{

const tn_tt = Number(r.tn_tt || 0);
const tn_tp = Number(r.tn_tp || 0);

const gq_tt = Number(r.gq_tt || 0);
const gq_tp = Number(r.gq_tp || 0);

const dx_tt = Number(r.dx_tt || 0);
const dx_tp = Number(r.dx_tp || 0);

const sohoa = Number(r.sohoa || 0);
const lamsach = Number(r.lam_sach || 0);
const chualamsach = Number(r.chualamsach || 0);

const qua_han = Number(r.qua_han || 0);
const tra_kq = Number(r.tra_kq || 0);
const thanhtoan = Number(r.thanhtoan || 0);

const tn = tn_tt + tn_tp;
const gq = gq_tt + gq_tp;
const sohoa_total = sohoa + lamsach + chualamsach;

const newRow = sheet.addRow([
r.stt || '',
r.ten_chitieu || '',

tn_tt,
tn_tp,
tn,

gq_tt,
gq_tp,
gq,

dx_tt,
dx_tp,

qua_han,
tra_kq,
thanhtoan,

sohoa,
lamsach,
chualamsach,
sohoa_total
]);

if(/^[A-Z]$/.test(r.stt)){
newRow.eachCell(c=>{
c.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'FFE4E4'}
};
});
}

if(/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)){
newRow.eachCell(c=>{
c.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'E8F5E9'}
};
});
}

/* ===== STYLE CELL ===== */

newRow.eachCell((cell, colNumber)=>{

if(colNumber >= 3){
cell.numFmt = '#,##0';
cell.alignment = { horizontal:'center' };
}else{
cell.alignment = { horizontal:'left' };
}

cell.border = {
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};

});

newRow.eachCell((cell, colNumber)=>{
if(colNumber >= 3){
cell.numFmt = '#,##0';
}
});

});

/* ===== TỔNG TOÀN TỈNH ===== */

let sum_tn_tt=0;
let sum_tn_tp=0;

let sum_gq_tt=0;
let sum_gq_tp=0;

let sum_dx_tt=0;
let sum_dx_tp=0;

let sum_quahan=0;
let sum_trakq=0;
let sum_thanhtoan=0;

let sum_sohoa=0;
let sum_lamsach=0;
let sum_chualamsach=0;

/* 🔥 TÍNH LẠI TỔNG */
for(let i=0;i<data.length;i++){

const r = data[i];

sum_tn_tt += Number(r.tn_tt||0);
sum_tn_tp += Number(r.tn_tp||0);

sum_gq_tt += Number(r.gq_tt||0);
sum_gq_tp += Number(r.gq_tp||0);

sum_dx_tt += Number(r.dx_tt||0);
sum_dx_tp += Number(r.dx_tp||0);

sum_quahan += Number(r.qua_han||0);
sum_trakq += Number(r.tra_kq||0);

sum_thanhtoan += Number(r.thanhtoan||0);

sum_sohoa += Number(r.sohoa||0);
sum_lamsach += Number(r.lam_sach||0);
sum_chualamsach += Number(r.chualamsach||0);

}

const sum_tn = sum_tn_tt + sum_tn_tp;
const sum_gq = sum_gq_tt + sum_gq_tp;
const sum_sohoa_total = sum_sohoa + sum_lamsach + sum_chualamsach;

/* 🔥 THÊM DÒNG */
const totalRow = sheet.addRow([
'',
'TỔNG TOÀN TỈNH',

sum_tn_tt,
sum_tn_tp,
sum_tn,

sum_gq_tt,
sum_gq_tp,
sum_gq,

sum_dx_tt,
sum_dx_tp,

sum_quahan,
sum_trakq,
sum_thanhtoan,

sum_sohoa,
sum_lamsach,
sum_chualamsach,
sum_sohoa_total
]);

/* STYLE */
totalRow.eachCell((cell, colNumber)=>{

cell.font = { bold:true };

cell.border = {
top:{style:'medium'},
left:{style:'thin'},
bottom:{style:'medium'},
right:{style:'thin'}
};

if(colNumber >= 3){
cell.numFmt = '#,##0';
cell.alignment = { horizontal:'center' };
}else{
cell.alignment = { horizontal:'left' };
}

cell.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'C3E6CB'}
};

});

sheet.columns = [
{width:6},
{width:40},
{width:10},{width:10},{width:12},
{width:10},{width:10},{width:12},
{width:10},{width:10},
{width:10},{width:12},{width:12},
{width:12},{width:12},{width:12},{width:14}
];

/* RESPONSE */
res.setHeader(
'Content-Type',
'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);

res.setHeader(
'Content-Disposition',
`attachment; filename=TongHop_Thang_${thang}.xlsx`
);

await workbook.xlsx.write(res);
res.end();

}catch(err){

console.error("Lỗi export:", err);

/* 🔥 chỉ trả 1 lần */
if(!res.headersSent){
res.status(500).json({error: err.message});
}

}

});

app.get("/test", async(req,res)=>{
const data = await pool.query("SELECT * FROM chitieu");
res.json(data.rows);
});

app.post("/login", async(req,res)=>{
const {username,password} = req.body;
const result = await pool.query(
"SELECT * FROM users WHERE username=$1 AND password=$2",
[username,password]
);
if(result.rows.length===0){
return res.json({success:false});
}
const user = result.rows[0];
res.json({
success:true,
role:user.role,
xa_id:user.xa_id
});
});

app.post("/mo-lai-baocao", async(req,res)=>{
try{
const {xa_id, thang} = req.body;
await pool.query(
`
UPDATE dulieu_baocao
SET trang_thai='draft'
WHERE xa_id=$1 AND thang=$2
`,
[xa_id, thang]
);
res.json({success:true});
}catch(err){
console.error(err);
res.status(500).send("Lỗi mở lại báo cáo");
}
});

app.get("/ten-xa/:id", async(req,res)=>{
const id = req.params.id;
const result = await pool.query(
"SELECT ten_xa FROM xa WHERE id=$1", [id]
);
res.json(result.rows[0]);
});

app.get("/kpi", async(req,res)=>{
const thang=req.query.thang;
const result=await pool.query(`
SELECT
SUM(tiepnhan_tructuyen+tiepnhan_tructiep) tiepnhan,
SUM(giaiquyet_tructuyen+giaiquyet_tructiep) giaiquyet,
SUM(dangxuly_tructuyen+dangxuly_tructiep) dangxuly
FROM dulieu_baocao
WHERE thang=$1
`,[thang]);
res.json(result.rows[0]);
});

app.get("/top-xa", async(req,res)=>{
const thang=req.query.thang;
const result=await pool.query(`
SELECT x.ten_xa,
SUM(d.tiepnhan_tructuyen+d.tiepnhan_tructiep) tong
FROM dulieu_baocao d
JOIN xa x ON x.id=d.xa_id
WHERE thang=$1
GROUP BY x.ten_xa
ORDER BY tong DESC
LIMIT 10
`,[thang]);
res.json(result.rows);
});

app.get("/trangthai-baocao", async(req,res)=>{
const {xa_id,thang}=req.query;
const result=await pool.query(
`SELECT trang_thai FROM dulieu_baocao WHERE xa_id=$1 AND thang=$2 LIMIT 1`, [xa_id,thang]
);
if(result.rows.length===0){
return res.json({trang_thai:"draft"});
}
res.json({trang_thai:result.rows[0].trang_thai});
});

app.post("/gui-baocao", async(req,res)=>{
const {xa_id,thang}=req.body;
await pool.query(
`UPDATE dulieu_baocao SET trang_thai='submit' WHERE xa_id=$1 AND thang=$2`, [xa_id,thang]
);
res.json({success:true});
await pool.query(
"INSERT INTO log_hoatdong(xa_id,hanh_dong) VALUES($1,$2)",
[xa_id,"Gửi báo cáo"]
);
});

app.post("/doi-mat-khau", async (req,res)=>{

try{

const {xa_id, oldPassword, newPassword} = req.body;

if(!xa_id || !oldPassword || !newPassword){
return res.json({success:false, message:"Thiếu dữ liệu"});
}

/* kiểm tra mật khẩu cũ */

const check = await pool.query(
"SELECT * FROM users WHERE xa_id=$1 AND password=$2",
[xa_id, oldPassword]
);

if(check.rows.length === 0){
return res.json({success:false});
}

/* cập nhật mật khẩu */

await pool.query(
"UPDATE users SET password=$1 WHERE xa_id=$2",
[newPassword, xa_id]
);

res.json({success:true});

}catch(err){

console.error("Lỗi đổi mật khẩu:", err);

/* ⚠️ TRẢ JSON để tránh crash frontend */
res.json({success:false, error:"server"});

}

});

app.get("/log", async(req,res)=>{
const result = await pool.query(`
SELECT l.*, x.ten_xa
FROM log_hoatdong l
LEFT JOIN xa x ON l.xa_id = x.id
ORDER BY l.thoi_gian DESC
LIMIT 100
`);
res.json(result.rows);
});

app.get("/export-xa", async (req,res)=>{

try{

const thang = req.query.thang;
const ten_xa = req.query.ten_xa;

/* query dữ liệu theo xã */
const result = await pool.query(`
SELECT 
c.stt,
c.ten_chitieu,

COALESCE(d.tiepnhan_tructuyen,0) as tn_tt,
COALESCE(d.tiepnhan_tructiep,0) as tn_tp,

COALESCE(d.giaiquyet_tructuyen,0) as gq_tt,
COALESCE(d.giaiquyet_tructiep,0) as gq_tp,

COALESCE(d.dangxuly_tructuyen,0) as dx_tt,
COALESCE(d.dangxuly_tructiep,0) as dx_tp,

COALESCE(d.ho_so_qua_han,0) as qua_han,
COALESCE(d.tra_kq_tructuyen,0) as tra_kq,

COALESCE(d.thanhtoan_tructuyen,0) as thanhtoan,

COALESCE(d.sohoa,0) as sohoa,
COALESCE(d.lam_sach,0) as lam_sach,
COALESCE(d.chua_lam_sach,0) as chualamsach

FROM chitieu c
LEFT JOIN dulieu_baocao d 
ON c.id = d.chitieu_id 
AND d.thang = $1

LEFT JOIN xa x ON d.xa_id = x.id

WHERE x.ten_xa ILIKE $2

ORDER BY c.id
`,[thang, `%${ten_xa}%`]);

const data = result.rows;

/* Excel */
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('BaoCaoXa');

/* tiêu đề */
sheet.addRow([`BÁO CÁO ĐƠN VỊ: ${ten_xa.toUpperCase()}`]);
sheet.mergeCells('A1:Q1');

sheet.getCell('A1').font = {bold:true, size:14};
sheet.getCell('A1').alignment = {horizontal:'center'};

/* ===== HEADER 2 DÒNG ===== */

sheet.mergeCells('A2:A3');
sheet.getCell('A2').value = 'STT';

sheet.mergeCells('B2:B3');
sheet.getCell('B2').value = 'Dịch vụ công';

sheet.mergeCells('C2:E2');
sheet.getCell('C2').value = 'HỒ SƠ TIẾP NHẬN';

sheet.mergeCells('F2:H2');
sheet.getCell('F2').value = 'ĐÃ GIẢI QUYẾT';

sheet.mergeCells('I2:J2');
sheet.getCell('I2').value = 'ĐANG GIẢI QUYẾT';

sheet.mergeCells('K2:K3');
sheet.getCell('K2').value = 'QUÁ HẠN';

sheet.mergeCells('L2:L3');
sheet.getCell('L2').value = 'TRẢ Kq';

sheet.mergeCells('M2:M3');
sheet.getCell('M2').value = 'THANH TOÁN';

sheet.mergeCells('N2:Q2');
sheet.getCell('N2').value = 'SỐ HOÁ';

/* dòng 3 */
sheet.getRow(3).values = [
'', '',
'Trực tuyến','Trực tiếp','Tổng',
'Trực tuyến','Trực tiếp','Tổng',
'Trực tuyến','Trực tiếp',
'', '', '',
'Đã SH','Làm sạch','Chưa sạch','Tổng'
];

[2,3].forEach(r=>{
sheet.getRow(r).eachCell(cell=>{
cell.font = {bold:true};
cell.alignment = {horizontal:'center', vertical:'middle'};

cell.fill = {
type:'pattern',
pattern:'solid',
fgColor:{argb:'D9EEF3'}
};

cell.border = {
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};

});
});

/* data */
let sum_tn_tt=0, sum_tn_tp=0;
let sum_gq_tt=0, sum_gq_tp=0;
let sum_dx_tt=0, sum_dx_tp=0;
let sum_quahan=0, sum_trakq=0, sum_thanhtoan=0;
let sum_sohoa=0, sum_lamsach=0, sum_chualamsach=0;

data.forEach(r=>{

const tn_tt = Number(r.tn_tt||0);
const tn_tp = Number(r.tn_tp||0);

const gq_tt = Number(r.gq_tt||0);
const gq_tp = Number(r.gq_tp||0);

const dx_tt = Number(r.dx_tt||0);
const dx_tp = Number(r.dx_tp||0);

const sohoa = Number(r.sohoa||0);
const lamsach = Number(r.lam_sach||0);
const chualamsach = Number(r.chualamsach||0);

const qua_han = Number(r.qua_han||0);
const tra_kq = Number(r.tra_kq||0);
const thanhtoan = Number(r.thanhtoan||0);

const tn = tn_tt + tn_tp;
const gq = gq_tt + gq_tp;
const sohoa_total = sohoa + lamsach + chualamsach;

/* cộng tổng */
sum_tn_tt += tn_tt;
sum_tn_tp += tn_tp;
sum_gq_tt += gq_tt;
sum_gq_tp += gq_tp;
sum_dx_tt += dx_tt;
sum_dx_tp += dx_tp;
sum_quahan += qua_han;
sum_trakq += tra_kq;
sum_thanhtoan += thanhtoan;
sum_sohoa += sohoa;
sum_lamsach += lamsach;
sum_chualamsach += chualamsach;

/* thêm dòng */
const row = sheet.addRow([
r.stt,
r.ten_chitieu,

tn_tt, tn_tp, tn,
gq_tt, gq_tp, gq,
dx_tt, dx_tp,

qua_han, tra_kq, thanhtoan,

sohoa, lamsach, chualamsach, sohoa_total
]);


/* style */
row.eachCell((cell,col)=>{
if(col>=3){
cell.numFmt='#,##0';
cell.alignment={horizontal:'center'};
}else{
cell.alignment={horizontal:'left'};
}

cell.border={
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};
});

/* màu */
if(/^[A-Z]$/.test(r.stt)){
row.eachCell(c=>c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE4E4'}});
}
if(/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)){
row.eachCell(c=>c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'E8F5E9'}});
}

});

/* ===== TÍNH TỔNG ===== */

const sum_tn = sum_tn_tt + sum_tn_tp;
const sum_gq = sum_gq_tt + sum_gq_tp;
const sum_sohoa_total = sum_sohoa + sum_lamsach + sum_chualamsach;

/* ===== DÒNG TỔNG ===== */

const totalRow = sheet.addRow([
'',
'TỔNG TOÀN ĐƠN VỊ',

sum_tn_tt, sum_tn_tp, sum_tn,
sum_gq_tt, sum_gq_tp, sum_gq,
sum_dx_tt, sum_dx_tp,

sum_quahan, sum_trakq, sum_thanhtoan,

sum_sohoa, sum_lamsach, sum_chualamsach, sum_sohoa_total
]);

/* STYLE TỔNG */
totalRow.eachCell((cell,col)=>{

cell.font={bold:true};

cell.fill={
type:'pattern',
pattern:'solid',
fgColor:{argb:'C3E6CB'}
};

cell.border={
top:{style:'medium'},
left:{style:'thin'},
bottom:{style:'medium'},
right:{style:'thin'}
};

if(col>=3){
cell.numFmt='#,##0';
cell.alignment={horizontal:'center'};
}else{
cell.alignment={horizontal:'left'};
}

});

sheet.columns = [
{width:6},{width:40},
{width:10},{width:10},{width:12},
{width:10},{width:10},{width:12},
{width:10},{width:10},
{width:10},{width:12},{width:12},
{width:12},{width:12},{width:12},{width:14}
];


/* download */
res.setHeader(
'Content-Type',
'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);

res.setHeader(
'Content-Disposition',
`attachment; filename=BaoCao_${ten_xa}.xlsx`
);

await workbook.xlsx.write(res);
res.end();

}catch(err){

console.error(err);
res.status(500).send("Lỗi export xã");

}

});

app.get("/export-tong-xa", async (req,res)=>{
try{

const thang = req.query.thang;

/* ===== QUERY GROUP BY XÃ ===== */
const result = await pool.query(`
SELECT 
x.id,
x.ten_xa,

SUM(COALESCE(d.tiepnhan_tructuyen,0)) as tn_tt,
SUM(COALESCE(d.tiepnhan_tructiep,0)) as tn_tp,

SUM(COALESCE(d.giaiquyet_tructuyen,0)) as gq_tt,
SUM(COALESCE(d.giaiquyet_tructiep,0)) as gq_tp,

SUM(COALESCE(d.dangxuly_tructuyen,0)) as dx_tt,
SUM(COALESCE(d.dangxuly_tructiep,0)) as dx_tp,

SUM(COALESCE(d.ho_so_qua_han,0)) as qua_han,
SUM(COALESCE(d.tra_kq_tructuyen,0)) as tra_kq,

SUM(COALESCE(d.thanhtoan_tructuyen,0)) as thanhtoan,

SUM(COALESCE(d.sohoa,0)) as sohoa,
SUM(COALESCE(d.lam_sach,0)) as lam_sach,
SUM(COALESCE(d.chua_lam_sach,0)) as chualamsach

FROM xa x
LEFT JOIN dulieu_baocao d 
ON x.id = d.xa_id AND d.thang = $1

GROUP BY x.id, x.ten_xa
ORDER BY x.id
`,[thang]);

const data = result.rows;

/* ===== EXCEL ===== */
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('TongHopCAT');

/* ===== HEADER ===== */

sheet.mergeCells('A1:A2');
sheet.getCell('A1').value = 'STT';

sheet.mergeCells('B1:B2');
sheet.getCell('B1').value = 'Tên đơn vị';

sheet.mergeCells('C1:E1');
sheet.getCell('C1').value = 'HỒ SƠ TIẾP NHẬN';

sheet.mergeCells('F1:H1');
sheet.getCell('F1').value = 'ĐÃ GIẢI QUYẾT';

sheet.mergeCells('I1:J1');
sheet.getCell('I1').value = 'ĐANG GIẢI QUYẾT';

sheet.mergeCells('K1:K2');
sheet.getCell('K1').value = 'QUÁ HẠN';

sheet.mergeCells('L1:L2');
sheet.getCell('L1').value = 'TRẢ KQ';

sheet.mergeCells('M1:M2');
sheet.getCell('M1').value = 'THANH TOÁN';

sheet.mergeCells('N1:Q1');
sheet.getCell('N1').value = 'SỐ HOÁ';

/* dòng 2 */
sheet.getRow(2).values = [
'', '',
'Trực tuyến','Trực tiếp','Tổng',
'Trực tuyến','Trực tiếp','Tổng',
'Trực tuyến','Trực tiếp',
'', '', '',
'Đã Số hoá','Làm sạch','Chưa sạch','Tổng'
];

/* style header */
[1,2].forEach(r=>{
sheet.getRow(r).eachCell(cell=>{
cell.font={bold:true};
cell.alignment={horizontal:'center', vertical:'middle'};
cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'D9EEF3'}};
cell.border={
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};
});
});

/* ===== DATA ===== */

let stt = 1;

/* tổng toàn tỉnh */
let sum_tn_tt=0, sum_tn_tp=0;
let sum_gq_tt=0, sum_gq_tp=0;
let sum_dx_tt=0, sum_dx_tp=0;
let sum_quahan=0, sum_trakq=0, sum_thanhtoan=0;
let sum_sohoa=0, sum_lamsach=0, sum_chualamsach=0;

data.forEach(r=>{

const tn_tt = Number(r.tn_tt||0);
const tn_tp = Number(r.tn_tp||0);
const gq_tt = Number(r.gq_tt||0);
const gq_tp = Number(r.gq_tp||0);
const dx_tt = Number(r.dx_tt||0);
const dx_tp = Number(r.dx_tp||0);

const sohoa = Number(r.sohoa||0);
const lamsach = Number(r.lam_sach||0);
const chualamsach = Number(r.chualamsach||0);

const qua_han = Number(r.qua_han||0);
const tra_kq = Number(r.tra_kq||0);
const thanhtoan = Number(r.thanhtoan||0);

const tn = tn_tt + tn_tp;
const gq = gq_tt + gq_tp;
const sohoa_total = sohoa + lamsach + chualamsach;

/* cộng tổng tỉnh */
sum_tn_tt += tn_tt;
sum_tn_tp += tn_tp;
sum_gq_tt += gq_tt;
sum_gq_tp += gq_tp;
sum_dx_tt += dx_tt;
sum_dx_tp += dx_tp;
sum_quahan += qua_han;
sum_trakq += tra_kq;
sum_thanhtoan += thanhtoan;
sum_sohoa += sohoa;
sum_lamsach += lamsach;
sum_chualamsach += chualamsach;

/* thêm dòng */
const row = sheet.addRow([
stt++,
r.ten_xa,

tn_tt, tn_tp, tn,
gq_tt, gq_tp, gq,
dx_tt, dx_tp,

qua_han, tra_kq, thanhtoan,

sohoa, lamsach, chualamsach, sohoa_total
]);

row.eachCell((cell,col)=>{
if(col>=3){
cell.numFmt='#,##0';
cell.alignment={horizontal:'center'};
}else{
cell.alignment={horizontal:'left'};
}
cell.border={
top:{style:'thin'},
left:{style:'thin'},
bottom:{style:'thin'},
right:{style:'thin'}
};
});

});

/* ===== TỔNG TOÀN TỈNH ===== */

const sum_tn = sum_tn_tt + sum_tn_tp;
const sum_gq = sum_gq_tt + sum_gq_tp;
const sum_sohoa_total = sum_sohoa + sum_lamsach + sum_chualamsach;

const totalRow = sheet.addRow([
'',
'TỔNG TOÀN TỈNH',

sum_tn_tt, sum_tn_tp, sum_tn,
sum_gq_tt, sum_gq_tp, sum_gq,
sum_dx_tt, sum_dx_tp,

sum_quahan, sum_trakq, sum_thanhtoan,

sum_sohoa, sum_lamsach, sum_chualamsach, sum_sohoa_total
]);

totalRow.eachCell((cell,col)=>{
cell.font={bold:true};
cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'C3E6CB'}};
cell.border={
top:{style:'medium'},
left:{style:'thin'},
bottom:{style:'medium'},
right:{style:'thin'}
};
if(col>=3){
cell.numFmt='#,##0';
cell.alignment={horizontal:'center'};
}
});

/* width */
sheet.columns = [
{width:6},{width:30},
{width:10},{width:10},{width:12},
{width:10},{width:10},{width:12},
{width:10},{width:10},
{width:10},{width:12},{width:12},
{width:12},{width:12},{width:12},{width:14}
];

/* download */
res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
res.setHeader('Content-Disposition',`attachment; filename=TongHopCAT.xlsx`);

await workbook.xlsx.write(res);
res.end();

}catch(err){
console.error(err);
res.status(500).send("Lỗi export tổng xã");
}
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server chạy port", PORT);
}); 