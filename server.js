const express = require("express");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const app = express();
const multer = require('multer');
const path = require('path');
//upload van ban
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    const xa_id = req.body.xa_id || "unknown";
    cb(null, `xa_${xa_id}_${Date.now()}_${file.originalname}`);
  }
});
const upload = multer({ storage: storage });

const JWT_SECRET = process.env.JWT_SECRET || "baocao_secret_key_2024";

app.use(bodyParser.json());
app.use(express.static("public"));

/* Theo dõi đăng nhập sai - khóa 5 phút sau 5 lần sai */
const loginAttempts = new Map();
const MAX_ATTEMPTS = 5;
const LOCK_TIME = 5 * 60 * 1000; // 5 phút

/* Middleware xác thực JWT */
// VERSION 2.0
// TRUNG TÂM QUẢN LÝ KỲ BÁO CÁO
// Mọi API mới của Version 2.0 phải sử dụng:
//   - getKyHienTai()
//   - req.user.xa_id
// Không sử dụng:
//   - req.query.thang
//   - req.body.xa_id
// Lấy kỳ báo cáo đang mở
async function getKyHienTai(client = pool) {
  const result = await client.query(`SELECT id, thang, nam, trang_thai FROM ky_baocao WHERE trang_thai = 'mo'
        ORDER BY id DESC LIMIT 1`);
  if (result.rows.length === 0) {
    throw new Error("Chưa có kỳ báo cáo đang mở.");
  }
  return result.rows[0];
}

async function getTongHopData(client = pool, ky_id) {
  const result = await client.query(`
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
ON c.id=d.chitieu_id
AND d.ky_id=$1
AND d.trang_thai='submit'
GROUP BY c.id,c.stt,c.ten_chitieu
ORDER BY c.id
`, [ky_id]);
  return result.rows;
}

// Kiểm tra kỳ báo cáo còn mở hay đã khóa
function kyDangMo(ky) {
  return ky.trang_thai === "mo";
}
function verifyToken(role) {
  return (req, res, next) => {
    const authHeader = req.headers.authorization;
    if (!authHeader) return res.status(401).json({ error: "Chưa đăng nhập" });
    const token = authHeader.split(" ")[1];
    try {
      const decoded = jwt.verify(token, JWT_SECRET);
      req.user = decoded;
      if (role && decoded.role !== role) {
        return res.status(403).json({ error: "Không có quyền" });
      }
      next();
    } catch (err) {
      return res.status(401).json({ error: "Token không hợp lệ" });
    }
  };
}

/* Kết nối database */

const { Pool } = require('pg');
const pool = new Pool({
  connectionString: process.env.DATABASE_URL || "postgres://postgres@db:5432/baocao",
  ssl: false
});

app.post('/upload-vanban', upload.single('file'), async (req, res) => {
  
  try {
    // ❌ check file trước
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: "Không có file"
      });
    }

    const { xa_id } = req.body;

    if (!xa_id) {
      return res.status(400).json({
        success: false,
        message: "Thiếu xa_id"
      });
    }

    // 🔎 lấy kỳ báo cáo
    const ky = await pool.query(
      "SELECT thang, nam FROM ky_baocao WHERE trang_thai='mo' LIMIT 1"
    );

    if (ky.rows.length === 0) {
      return res.status(400).json({
        success: false,
        message: "Chưa mở kỳ báo cáo"
      });
    }

    const { thang, nam } = ky.rows[0];

    // 🧹 xóa bản cũ
    await pool.query(
      "DELETE FROM vanban WHERE xa_id=$1 AND thang=$2 AND nam=$3",
      [xa_id, thang, nam]
    );

    // 💾 insert mới
    await pool.query(
      `INSERT INTO vanban(xa_id, filename, original_name, thang, nam)
       VALUES($1,$2,$3,$4,$5)`,
      [xa_id, req.file.filename, req.file.originalname, thang, nam]
    );

    // ✅ CHỈ trả 1 lần duy nhất
    return res.json({
      success: true,
      message: "Upload thành công"
    });

  } catch (err) {
    console.error("Upload lỗi:", err);

    return res.status(500).json({
      success: false,
      message: "Lỗi server"
    });
  }
});

app.use('/uploads', require('express').static('uploads'));

app.get('/vanban', async (req, res) => {
  try {
    const { thang } = req.query;

    const result = await pool.query(`
      SELECT v.*, x.ten_xa
      FROM vanban v
      JOIN xa x ON v.xa_id = x.id
      WHERE v.thang = $1
      ORDER BY v.created_at DESC
    `, [thang]);

    res.json(result.rows);

  } catch (err) {
    console.error(err);
    res.status(500).json([]);
  }
});

app.get("/ky-baocao", async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT
          id,
          thang,
          nam,
          ngay_bat_dau,
          ngay_ket_thuc,
          trang_thai
      FROM ky_baocao
      WHERE trang_thai='mo'
      ORDER BY id DESC
      LIMIT 1
    `);
    if (result.rows.length === 0) {
      return res.status(404).json({
        success: false,
        message: "Chưa có kỳ báo cáo đang mở"
      });
    }
    res.json({
      success: true,
      ...result.rows[0]
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: "Lỗi server"
    });
  }
});
/* DANH SÁCH KỲ BÁO CÁO */
app.get("/ky-baocao-list", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT
          id,
          thang,
          nam,
          ngay_bat_dau,
          ngay_ket_thuc,
          trang_thai
      FROM ky_baocao
      ORDER BY nam DESC, thang DESC
    `);
    res.json({
      success: true,
      data: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: "Lỗi server"
    });
  }
});

app.get("/nhom-tthc", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT DISTINCT ma_nhom
      FROM xa
      WHERE ma_nhom IS NOT NULL
      ORDER BY ma_nhom
    `);
    res.json({
      success: true,
      data: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.get("/admin/chitieu", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT
          id,
          stt,
          ten_chitieu,
          cap_dong
      FROM chitieu
      WHERE stt <> 'TT'
      ORDER BY id
    `);
    res.json({
      success: true,
      data: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.post("/save-phanquyen", verifyToken("admin"), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    await client.query("DELETE FROM phanquyen_tthc");
    for (const item of req.body.data) {
      await client.query(`INSERT INTO phanquyen_tthc (ma_nhom, chitieu_id) VALUES ($1,$2)`,
        [item.ma_nhom, item.chitieu_id]);
    }
    await client.query("COMMIT");
    res.json({
      success: true
    });
  } catch (err) {
    await client.query("ROLLBACK");
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  } finally {
    client.release();
  }
});

app.get("/phanquyen-tthc", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT
          ma_nhom,
          chitieu_id
      FROM phanquyen_tthc
      ORDER BY ma_nhom,chitieu_id
    `);
    res.json({
      success: true,
      data: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.post("/phanquyen-tthc", verifyToken("admin"), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const data = req.body;
    await client.query("DELETE FROM phanquyen_tthc");
    for (const row of data) {
      await client.query(
        `
        INSERT INTO phanquyen_tthc(ma_nhom,chitieu_id)
        VALUES($1,$2)
        `,
        [
          row.ma_nhom,
          row.chitieu_id
        ]
      );
    }
    await client.query("COMMIT");
    res.json({
      success: true,
      message: "Đã lưu phân quyền"
    });
  } catch (err) {
    await client.query("ROLLBACK");
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  } finally {
    client.release();
  }
});

app.post("/chuyen-ky-baocao", verifyToken("admin"), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const { thang, nam, ngay_bat_dau, ngay_ket_thuc } = req.body;
    // 1. Lấy kỳ hiện tại
    const ky = await getKyHienTai(client);
    // 2. Không cho tạo trùng
    const check = await client.query(`SELECT id FROM ky_baocao WHERE thang=$1 AND nam=$2`, [thang, nam]);
    if (check.rows.length > 0) {
      throw new Error("Kỳ báo cáo này đã tồn tại.");
    }
    // 3. Không cho quay lùi
    if (
      nam < ky.nam || (nam == ky.nam && thang <= ky.thang)
    ) {
      throw new Error("Kỳ mới phải lớn hơn kỳ hiện tại.");
    }
    // 4. Đóng kỳ cũ
    await client.query(`UPDATE ky_baocao SET trang_thai='dong' WHERE id=$1`, [ky.id]);
    // 5. Tạo kỳ mới
    const insert = await client.query(`INSERT INTO ky_baocao(thang, nam, ngay_bat_dau, ngay_ket_thuc, trang_thai)
      VALUES ($1, $2, $3, $4, 'mo') RETURNING id`, [thang, nam, ngay_bat_dau, ngay_ket_thuc]);
    // 6. Ghi log
    await client.query(
      `INSERT INTO log_hoatdong (xa_id, hanh_dong) VALUES (NULL, $1)`,
      [`Chuyển kỳ báo cáo sang Tháng ${thang}/${nam}`]
    );
    await client.query("COMMIT");
    res.json({
      success: true, message: "Đã chuyển kỳ báo cáo.", id: insert.rows[0].id
    });
  }
  catch (err) {
    await client.query("ROLLBACK");
    console.error(err);
    res.status(400).json({
      success: false, message: err.message
    });
  }
  finally {
    client.release();
  }
});

app.get("/dashboard-xa", verifyToken("admin"), async (req, res) => {
  try {
    const ky = await getKyHienTai();
    const result = await pool.query(`SELECT x.id, x.ten_xa, x.ma_nhom, MAX(d.updated_at) AS cap_nhat_cuoi,
    CASE
        WHEN COUNT(d.id)=0 THEN 'Chưa nhập'
        WHEN BOOL_OR(d.trang_thai='submit') THEN 'Đã gửi'
        ELSE 'Đang nhập'
    END AS trang_thai FROM xa x LEFT JOIN dulieu_baocao d ON x.id=d.xa_id AND d.ky_id=$1
    GROUP BY x.id, x.ten_xa, x.ma_nhom
    ORDER BY x.ten_xa`, [ky.id]);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

/*Xem log*/
app.get("/log", verifyToken("admin"), async (req, res) => {
  const result = await pool.query(`SELECT l.*,x.ten_xa FROM log_hoatdong l LEFT JOIN xa x ON l.xa_id=x.id
  ORDER BY l.thoi_gian DESC LIMIT 100`);
  res.json(result.rows);
});

/* reset mật khẩu*/
app.post("/reset-mat-khau", verifyToken("admin"), async (req, res) => {
  try {
    const { xa_id } = req.body;
    const hashed = await bcrypt.hash("123456", 10);
    await pool.query(
      "UPDATE users SET password=$1 WHERE xa_id=$2",
      [hashed, xa_id]
    );
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).send("Lỗi server");
  }
});

app.get("/chitieu", verifyToken(), async (req, res) => {
  try {
    // Admin và Viewer xem toàn bộ
    if (req.user.role === "admin" || req.user.role === "viewer") {
      const result = await pool.query(`
        SELECT
            id,
            stt,
            ten_chitieu,
            cap_dong
        FROM chitieu
        WHERE stt <> 'TT'
        ORDER BY id
      `);
      return res.json(result.rows);
    }
    // Lấy mã nhóm của đơn vị
    const xaResult = await pool.query(
      `
      SELECT ma_nhom
      FROM xa
      WHERE id=$1
      `,
      [req.user.xa_id]
    );
    if (xaResult.rows.length === 0) {
      return res.json([]);
    }
    const maNhom = xaResult.rows[0].ma_nhom;
    // Chỉ lấy TTHC được phân quyền
    const result = await pool.query(
      `
      SELECT
          c.id,
          c.stt,
          c.ten_chitieu,
          c.cap_dong
      FROM chitieu c
      INNER JOIN phanquyen_tthc p
          ON c.id = p.chitieu_id
      WHERE
          p.ma_nhom = $1
      AND c.stt <> 'TT'
      ORDER BY c.id
      `,
      [maNhom]
    );
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).send("Lỗi server");
  }
});

app.get("/dulieu-xa", verifyToken("xa"), async (req, res) => {
  try {
    // Lấy xã từ JWT
    const xa_id = req.user.xa_id;
    // Lấy kỳ đang mở
    const ky = await getKyHienTai();
    const result = await pool.query(`SELECT * FROM dulieu_baocao WHERE xa_id = $1 AND ky_id = $2 ORDER BY chitieu_id`,
      [xa_id, ky.id]
    );
    res.json({
      success: true,
      ky: { id: ky.id, thang: ky.thang, nam: ky.nam }, data: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get("/tonghop", async (req, res) => {
  try {
    const ky = await getKyHienTai();
    const ky_id = ky.id;
    const thang = ky.thang;
    const nam = ky.nam;
    const data = await getTongHopData(pool, ky_id);
    res.json(data);
  } catch (err) {
    console.error("Lỗi tonghop:", err);
    /* 🔥 QUAN TRỌNG: trả JSON */
    res.json([]);
  }
});

app.get("/export-excel", verifyToken("admin"), async (req, res) => {
  try {
    const ky = await getKyHienTai();
    const ky_id = ky.id;
    /* EXCEL */
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('TongHop');
    /* QUERY */
    const data = await getTongHopData(pool, ky_id);
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
      'TT', 'TP', 'Tổng',
      'TT', 'TP', 'Tổng',
      'TT', 'TP',
      '', '', '',
      'Đã SH', 'Làm sạch', 'Chưa sạch', 'Tổng'
    ];

    /* ===== STYLE HEADER ===== */

    [1, 2].forEach(rowIndex => {
      sheet.getRow(rowIndex).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EEF3' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* DATA */
    data.forEach(r => {
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
        tn_tt, tn_tp, tn, gq_tt, gq_tp, gq, dx_tt, dx_tp,
        qua_han, tra_kq, thanhtoan, sohoa, lamsach, chualamsach, sohoa_total
      ]);
      if (/^[A-Z]$/.test(r.stt)) {
        newRow.eachCell(c => {
          c.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE4E4' }
          };
        });
      }
      if (/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)) {
        newRow.eachCell(c => {
          c.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E8F5E9' }
          };
        });
      }
      /* ===== STYLE CELL ===== */
      newRow.eachCell((cell, colNumber) => {
        if (colNumber >= 3) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      newRow.eachCell((cell, colNumber) => {
        if (colNumber >= 3) {
          cell.numFmt = '#,##0';
        }
      });
    });
    /* ===== TỔNG TOÀN TỈNH ===== */
    let sum_tn_tt = 0;
    let sum_tn_tp = 0;
    let sum_gq_tt = 0;
    let sum_gq_tp = 0;
    let sum_dx_tt = 0;
    let sum_dx_tp = 0;
    let sum_quahan = 0;
    let sum_trakq = 0;
    let sum_thanhtoan = 0;
    let sum_sohoa = 0;
    let sum_lamsach = 0;
    let sum_chualamsach = 0;
    /* 🔥 TÍNH LẠI TỔNG */
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      sum_tn_tt += Number(r.tn_tt || 0);
      sum_tn_tp += Number(r.tn_tp || 0);
      sum_gq_tt += Number(r.gq_tt || 0);
      sum_gq_tp += Number(r.gq_tp || 0);
      sum_dx_tt += Number(r.dx_tt || 0);
      sum_dx_tp += Number(r.dx_tp || 0);
      sum_quahan += Number(r.qua_han || 0);
      sum_trakq += Number(r.tra_kq || 0);
      sum_thanhtoan += Number(r.thanhtoan || 0);
      sum_sohoa += Number(r.sohoa || 0);
      sum_lamsach += Number(r.lam_sach || 0);
      sum_chualamsach += Number(r.chualamsach || 0);
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
    totalRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      };
      if (colNumber >= 3) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'center' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C3E6CB' }
      };
    });
    sheet.columns = [
      { width: 6 },
      { width: 40 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 },
      { width: 10 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }
    ];
    /* RESPONSE */
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=TongHop_Thang_${ky.thang}_${ky.nam}.xlsx`
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Lỗi export:", err);
    /* 🔥 chỉ trả 1 lần */
    if (!res.headersSent) {
      res.status(500).json({ error: err.message });
    }
  }
});

app.get("/test", async (req, res) => {
  const data = await pool.query("SELECT * FROM chitieu");
  res.json(data.rows);
});

app.post("/login", async (req, res) => {
  const { username, password } = req.body;

  // Kiểm tra tài khoản có đang bị khóa không
  const attempt = loginAttempts.get(username);
  if (attempt && attempt.count >= MAX_ATTEMPTS) {
    const elapsed = Date.now() - attempt.lastTime;
    if (elapsed < LOCK_TIME) {
      const remaining = Math.ceil((LOCK_TIME - elapsed) / 60000);
      return res.json({ success: false, locked: true, message: `Tài khoản bị khóa. Vui lòng thử lại sau ${remaining} phút.` });
    }
    // Hết thời gian khóa, reset
    loginAttempts.delete(username);
  }

  const result = await pool.query(
    "SELECT * FROM users WHERE username=$1",
    [username]
  );
  if (result.rows.length === 0) {
    // Tăng số lần đăng nhập sai
    const cur = loginAttempts.get(username) || { count: 0, lastTime: 0 };
    cur.count++;
    cur.lastTime = Date.now();
    loginAttempts.set(username, cur);
    const left = MAX_ATTEMPTS - cur.count;
    if (left <= 0) {
      return res.json({ success: false, locked: true, message: "Đăng nhập sai quá 5 lần. Tài khoản bị khóa 5 phút." });
    }
    return res.json({ success: false, message: `Sai tên đăng nhập hoặc mật khẩu. Còn ${left} lần thử.` });
  }
  const user = result.rows[0];
  const match = await bcrypt.compare(password, user.password);
  if (!match) {
    const cur = loginAttempts.get(username) || { count: 0, lastTime: 0 };
    cur.count++;
    cur.lastTime = Date.now();
    loginAttempts.set(username, cur);
    const left = MAX_ATTEMPTS - cur.count;
    if (left <= 0) {
      return res.json({ success: false, locked: true, message: "Đăng nhập sai quá 5 lần. Tài khoản bị khóa 5 phút." });
    }
    return res.json({ success: false, message: `Sai tên đăng nhập hoặc mật khẩu. Còn ${left} lần thử.` });
  }

  // Đăng nhập thành công, xóa lịch sử đăng nhập sai
  loginAttempts.delete(username);

  const token = jwt.sign(
    { id: user.id, role: user.role, xa_id: user.xa_id },
    JWT_SECRET,
    { expiresIn: "8h" }
  );
  res.json({
    success: true,
    token: token,
    role: user.role,
    xa_id: user.xa_id
  });
});

app.post("/mo-lai-baocao", verifyToken("admin"), async (req, res) => {
  try {
    const { xa_id } = req.body;
    const ky = await getKyHienTai();
    await pool.query(`UPDATE dulieu_baocao SET trang_thai='draft' WHERE xa_id=$1 AND ky_id=$2`, [xa_id, ky.id]);
    await pool.query(`INSERT INTO log_hoatdong (xa_id,hanh_dong) VALUES($1,$2)`,
      [xa_id, `Admin mở lại báo cáo kỳ ${ky.thang}/${ky.nam}`]);
    res.json({
      success: true
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.get("/ten-xa/:id", async (req, res) => {
  const id = req.params.id;
  const result = await pool.query(
    "SELECT ten_xa FROM xa WHERE id=$1", [id]
  );
  res.json(result.rows[0]);
});

app.get("/trangthai-baocao", verifyToken("xa"), async (req, res) => {
  try {
    const xa_id = req.user.xa_id;
    const ky = await getKyHienTai();
    const result = await pool.query(`SELECT trang_thai FROM dulieu_baocao WHERE xa_id = $1 AND ky_id = $2 LIMIT 1`,
      [xa_id, ky.id]
    );
    if (result.rows.length === 0) {
      return res.json({
        success: true,
        trang_thai: "draft"
      });
    }
    res.json({
      success: true,
      trang_thai: result.rows[0].trang_thai
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.post("/luu", verifyToken("xa"), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const xa_id = req.user.xa_id;
    const ky = await getKyHienTai(client);
    const data = req.body;
    if (!Array.isArray(data) || data.length === 0) {
      await client.query("ROLLBACK");
      return res.json({
        success: true
      });
    }
    /* Xóa dữ liệu cũ */
    await client.query(`DELETE FROM dulieu_baocao WHERE xa_id=$1 AND ky_id=$2`, [xa_id, ky.id]);
    for (const row of data) {
      await client.query(
        `INSERT INTO dulieu_baocao(
          ky_id,
          xa_id,
          chitieu_id,
          tiepnhan_tructuyen,
          tiepnhan_tructiep,
          giaiquyet_tructuyen,
          giaiquyet_tructiep,
          dangxuly_tructuyen,
          dangxuly_tructiep,
          thanhtoan_tructuyen,
          sohoa,
          lam_sach,
          chua_lam_sach,
          ho_so_qua_han,
          tra_kq_tructuyen,
          trang_thai
        )
        VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,'draft')`,
        [
          ky.id,
          xa_id,
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
        ]
      );
    }
    /* ghi log */
    await client.query(
      `INSERT INTO log_hoatdong(xa_id,hanh_dong) VALUES($1,$2)`,
      [xa_id, `Lưu tạm báo cáo kỳ ${ky.thang}/${ky.nam}`]
    );
    await client.query("COMMIT");
    res.json({
      success: true,
      message: "Đã lưu tạm."
    });
  } catch (err) {
    await client.query("ROLLBACK");
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  } finally {
    client.release();
  }
});

app.post("/gui-baocao", verifyToken("xa"), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const xa_id = req.user.xa_id;
    const ky = await getKyHienTai(client);
    // Kiểm tra xã đã có dữ liệu hay chưa
    const check = await client.query(`SELECT COUNT(*)::int AS total FROM dulieu_baocao WHERE xa_id = $1 AND ky_id = $2`,
      [xa_id, ky.id]
    );
    if (check.rows[0].total === 0) {
      await client.query("ROLLBACK");
      return res.status(400).json({
        success: false,
        message: "Chưa có dữ liệu để gửi báo cáo."
      });
    }
    // Chuyển trạng thái sang submit
    await client.query(`UPDATE dulieu_baocao SET trang_thai = 'submit' WHERE xa_id = $1 AND ky_id = $2`, [xa_id, ky.id]);
    // Ghi log
    await client.query(`INSERT INTO log_hoatdong(xa_id, hanh_dong) VALUES($1,$2)`,
      [xa_id, `Gửi báo cáo kỳ ${ky.thang}/${ky.nam}`]
    );
    await client.query("COMMIT");
    res.json({
      success: true,
      message: "Đã gửi báo cáo thành công."
    });
  } catch (err) {
    await client.query("ROLLBACK");
    console.error(err);
    res.status(500).json({
      success: false,
      message: "Lỗi gửi báo cáo."
    });
  } finally {
    client.release();
  }
});

app.post("/doi-mat-khau", async (req, res) => {

  try {

    const { xa_id, oldPassword, newPassword } = req.body;

    if (!xa_id || !oldPassword || !newPassword) {
      return res.json({ success: false, message: "Thiếu dữ liệu" });
    }

    /* kiểm tra mật khẩu cũ */

    const check = await pool.query(
      "SELECT * FROM users WHERE xa_id=$1",
      [xa_id]
    );

    if (check.rows.length === 0) {
      return res.json({ success: false });
    }

    const match = await bcrypt.compare(oldPassword, check.rows[0].password);
    if (!match) {
      return res.json({ success: false });
    }

    /* cập nhật mật khẩu */

    const hashed = await bcrypt.hash(newPassword, 10);
    await pool.query(
      "UPDATE users SET password=$1 WHERE xa_id=$2",
      [hashed, xa_id]
    );

    res.json({ success: true });

  } catch (err) {

    console.error("Lỗi đổi mật khẩu:", err);

    /* ⚠️ TRẢ JSON để tránh crash frontend */
    res.json({ success: false, error: "server" });

  }

});

app.get("/export-xa", async (req, res) => {
  try {
    const thang = req.query.thang;
    const xa_id = req.query.xa_id;
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
WHERE d.xa_id = $2
ORDER BY c.id
`, [thang, xa_id]);
    const data = result.rows;
    if (!data || data.length === 0) {
      return res.status(404).send("Không có dữ liệu để export");
    }
    /* Excel */
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('BaoCaoXa');
    /* tiêu đề */
    const xaInfo = await pool.query(
      "SELECT ten_xa FROM xa WHERE id = $1",
      [xa_id]
    );
    const tenXa = xaInfo.rows[0]?.ten_xa || xa_id;
    sheet.addRow([`BÁO CÁO ĐƠN VỊ: ${tenXa.toUpperCase()}`]);
    sheet.mergeCells('A1:Q1');
    sheet.getCell('A1').font = { bold: true, size: 14 };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
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
    sheet.getCell('L2').value = 'TRẢ KQ';
    sheet.mergeCells('M2:M3');
    sheet.getCell('M2').value = 'THANH TOÁN';
    sheet.mergeCells('N2:Q2');
    sheet.getCell('N2').value = 'SỐ HOÁ';
    /* dòng 3 */
    sheet.getRow(3).values = [
      '', '',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp',
      '', '', '',
      'Đã SH', 'Làm sạch', 'Chưa sạch', 'Tổng'
    ];
    [2, 3].forEach(r => {
      sheet.getRow(r).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EEF3' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    /* data */
    let sum_tn_tt = 0, sum_tn_tp = 0;
    let sum_gq_tt = 0, sum_gq_tp = 0;
    let sum_dx_tt = 0, sum_dx_tp = 0;
    let sum_quahan = 0, sum_trakq = 0, sum_thanhtoan = 0;
    let sum_sohoa = 0, sum_lamsach = 0, sum_chualamsach = 0;
    data.forEach(r => {
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
      row.eachCell((cell, col) => {
        if (col >= 3) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      /* màu */
      if (/^[A-Z]$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4E4' } });
      }
      if (/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } });
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
    totalRow.eachCell((cell, col) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C3E6CB' }
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      };
      if (col >= 3) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'center' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    });
    sheet.columns = [
      { width: 6 }, { width: 40 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 },
      { width: 10 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }
    ];
    /* download */
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=BaoCao_${xa_id}.xlsx`
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Lỗi export xã");
  }
});

app.get("/export-donvi", verifyToken("xa"), async (req, res) => {
  try {
    const xa_id = req.user.xa_id;
    const ky = await getKyHienTai();
    const ky_id = ky.id;
    const thang = ky.thang;
    const nam = ky.nam;
    const check = await pool.query(`
SELECT trang_thai
FROM dulieu_baocao
WHERE xa_id=$1
AND ky_id=$2
LIMIT 1
`, [xa_id, ky.id]);
    if (
      check.rows.length === 0 ||
      check.rows[0].trang_thai !== "submit"
    ) {
      return res.status(400).send(
        "Bạn cần gửi báo cáo trước khi xuất Excel."
      );
    }

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
ON d.chitieu_id = c.id
AND d.xa_id = $2
AND d.ky_id = $1
ORDER BY c.id
`, [ky_id, xa_id]);
    const data = result.rows;
    if (!data || data.length === 0) {
      return res.status(404).send("Không có dữ liệu để export");
    }
    /* Excel */
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('BaoCaoDonVi');
    /* tiêu đề */
    const xaInfo = await pool.query(
      "SELECT ten_xa FROM xa WHERE id = $1",
      [xa_id]
    );
    const tenXa = xaInfo.rows[0]?.ten_xa || xa_id;
    const tenFile = tenXa.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").replace(/\s+/g, "");
    sheet.addRow([`BÁO CÁO ĐƠN VỊ: ${tenXa.toUpperCase()}`]);
    sheet.mergeCells('A1:Q1');
    sheet.getCell('A1').font = { bold: true, size: 14 };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
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
    sheet.getCell('L2').value = 'TRẢ KQ';
    sheet.mergeCells('M2:M3');
    sheet.getCell('M2').value = 'THANH TOÁN';
    sheet.mergeCells('N2:Q2');
    sheet.getCell('N2').value = 'SỐ HOÁ';
    /* dòng 3 */
    sheet.getRow(3).values = [
      '', '',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp',
      '', '', '',
      'Đã SH', 'Làm sạch', 'Chưa sạch', 'Tổng'
    ];
    [2, 3].forEach(r => {
      sheet.getRow(r).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EEF3' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    /* data */
    let sum_tn_tt = 0, sum_tn_tp = 0;
    let sum_gq_tt = 0, sum_gq_tp = 0;
    let sum_dx_tt = 0, sum_dx_tp = 0;
    let sum_quahan = 0, sum_trakq = 0, sum_thanhtoan = 0;
    let sum_sohoa = 0, sum_lamsach = 0, sum_chualamsach = 0;
    data.forEach(r => {
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
      row.eachCell((cell, col) => {
        if (col >= 3) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      /* màu */
      if (/^[A-Z]$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4E4' } });
      }
      if (/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } });
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
    totalRow.eachCell((cell, col) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C3E6CB' }
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      };
      if (col >= 3) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'center' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    });
    sheet.columns = [
      { width: 6 }, { width: 40 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 },
      { width: 10 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }
    ];
    /* download */
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=BaoCao_${tenFile}_Thang${thang}.xlsx`
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("===== EXPORT DONVI =====");
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.get("/export-donvi-admin", verifyToken("admin"), async (req, res) => {
  try {
    const xa_id = Number(req.query.id);
    const xaName = await pool.query("SELECT ten_xa FROM xa WHERE id=$1", [xa_id]);
    function removeVietnamese(str) {
      return str
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d")
        .replace(/Đ/g, "D")
        .replace(/\s+/g, "_");
    }
    const tenDonVi = removeVietnamese(xaName.rows[0].ten_xa);
    const ky = await getKyHienTai();
    const ky_id = ky.id;
    const thang = ky.thang;
    const nam = ky.nam;
    const check = await pool.query(`
SELECT trang_thai
FROM dulieu_baocao
WHERE xa_id=$1
AND ky_id=$2
LIMIT 1
`, [xa_id, ky.id]);
    if (
      check.rows.length === 0 ||
      check.rows[0].trang_thai !== "submit"
    ) {
      return res.status(400).send(
        "Bạn cần gửi báo cáo trước khi xuất Excel."
      );
    }

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
ON d.chitieu_id = c.id
AND d.xa_id = $2
AND d.ky_id = $1
ORDER BY c.id
`, [ky_id, xa_id]);
    const data = result.rows;
    if (!data || data.length === 0) {
      return res.status(404).send("Không có dữ liệu để export");
    }
    /* Excel */
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('BaoCaoDonVi');
    /* tiêu đề */
    const xaInfo = await pool.query(
      "SELECT ten_xa FROM xa WHERE id = $1",
      [xa_id]
    );
    const tenXa = xaInfo.rows[0]?.ten_xa || xa_id;
    const tenFile = tenXa.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").replace(/\s+/g, "");
    sheet.addRow([`BÁO CÁO ĐƠN VỊ: ${tenXa.toUpperCase()}`]);
    sheet.mergeCells('A1:Q1');
    sheet.getCell('A1').font = { bold: true, size: 14 };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
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
    sheet.getCell('L2').value = 'TRẢ KQ';
    sheet.mergeCells('M2:M3');
    sheet.getCell('M2').value = 'THANH TOÁN';
    sheet.mergeCells('N2:Q2');
    sheet.getCell('N2').value = 'SỐ HOÁ';
    /* dòng 3 */
    sheet.getRow(3).values = [
      '', '',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp',
      '', '', '',
      'Đã SH', 'Làm sạch', 'Chưa sạch', 'Tổng'
    ];
    [2, 3].forEach(r => {
      sheet.getRow(r).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EEF3' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    /* data */
    let sum_tn_tt = 0, sum_tn_tp = 0;
    let sum_gq_tt = 0, sum_gq_tp = 0;
    let sum_dx_tt = 0, sum_dx_tp = 0;
    let sum_quahan = 0, sum_trakq = 0, sum_thanhtoan = 0;
    let sum_sohoa = 0, sum_lamsach = 0, sum_chualamsach = 0;
    data.forEach(r => {
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
      row.eachCell((cell, col) => {
        if (col >= 3) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      /* màu */
      if (/^[A-Z]$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4E4' } });
      }
      if (/^(I|II|III|IV|V|VI|VII|VIII|IX)$/.test(r.stt)) {
        row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } });
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
    totalRow.eachCell((cell, col) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C3E6CB' }
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      };
      if (col >= 3) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'center' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    });
    sheet.columns = [
      { width: 6 }, { width: 40 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 },
      { width: 10 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }
    ];
    /* download */
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    const fileName =
      `BaoCao_${tenDonVi}_Thang${thang}_${nam}.xlsx`;

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${fileName}"`
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("===== EXPORT DONVI =====");
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message
    });
  }
});

app.get("/export-tong-donvi", verifyToken("admin"), async (req, res) => {
  try {
    const ky = await getKyHienTai();
    const ky_id = ky.id;
    const thang = ky.thang;
    const nam = ky.nam;
    const result = await pool.query(`
SELECT
    x.id,
    x.ten_xa,
    x.ma_nhom,
    COALESCE(SUM(d.tiepnhan_tructuyen),0) AS tn_tt,
    COALESCE(SUM(d.tiepnhan_tructiep),0) AS tn_tp,
    COALESCE(SUM(d.giaiquyet_tructuyen),0) AS gq_tt,
    COALESCE(SUM(d.giaiquyet_tructiep),0) AS gq_tp,
    COALESCE(SUM(d.dangxuly_tructuyen),0) AS dx_tt,
    COALESCE(SUM(d.dangxuly_tructiep),0) AS dx_tp,
    COALESCE(SUM(d.ho_so_qua_han),0) AS qua_han,
    COALESCE(SUM(d.tra_kq_tructuyen),0) AS tra_kq,
    COALESCE(SUM(d.thanhtoan_tructuyen),0) AS thanhtoan,
    COALESCE(SUM(d.sohoa),0) AS sohoa,
    COALESCE(SUM(d.lam_sach),0) AS lam_sach,
    COALESCE(SUM(d.chua_lam_sach),0) AS chualamsach

FROM xa x

LEFT JOIN dulieu_baocao d
ON d.xa_id = x.id
AND d.ky_id = $1
AND d.trang_thai='submit'

GROUP BY
x.id,
x.ten_xa,
x.ma_nhom

ORDER BY
x.id
`, [ky_id]);

    const data = result.rows;
    /* ===== EXCEL ===== */
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('TongHopDonVi');
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
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp', 'Tổng',
      'Trực tuyến', 'Trực tiếp',
      '', '', '',
      'Đã Số hoá', 'Làm sạch', 'Chưa sạch', 'Tổng'
    ];
    /* style header */
    [1, 2].forEach(r => {
      sheet.getRow(r).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EEF3' } };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    /* ===== DATA ===== */
    let stt = 1;
    /* tổng toàn tỉnh */
    let sum_tn_tt = 0, sum_tn_tp = 0;
    let sum_gq_tt = 0, sum_gq_tp = 0;
    let sum_dx_tt = 0, sum_dx_tp = 0;
    let sum_quahan = 0, sum_trakq = 0, sum_thanhtoan = 0;
    let sum_sohoa = 0, sum_lamsach = 0, sum_chualamsach = 0;
    data.forEach(r => {
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
      row.eachCell((cell, col) => {
        if (col >= 3) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
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
    totalRow.eachCell((cell, col) => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C3E6CB' } };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      };
      if (col >= 3) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'center' };
      }
    });
    /* width */
    sheet.columns = [
      { width: 6 }, { width: 30 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 10 },
      { width: 10 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }
    ];
    /* download */
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=TongHopDonVi_Thang${thang}_${nam}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Lỗi export tổng xã");
  }
});
/* ===== QUẢN LÝ TÀI KHOẢN ===== */
/* Danh sách tài khoản */
app.get("/users", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query(`
SELECT u.id, u.username, u.role, u.xa_id, x.ten_xa
FROM users u
LEFT JOIN xa x ON u.xa_id = x.id
ORDER BY u.id
`);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Lỗi server" });
  }
});

/* Danh sách xã (cho dropdown) */
app.get("/danh-sach-xa", verifyToken("admin"), async (req, res) => {
  try {
    const result = await pool.query("SELECT id, ten_xa FROM xa ORDER BY ten_xa");
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Lỗi server" });
  }
});

/* Thêm tài khoản */
app.post("/users", verifyToken("admin"), async (req, res) => {
  try {
    const { username, password, role, xa_id } = req.body;
    if (!username || !password || !role) {
      return res.json({ success: false, message: "Thiếu dữ liệu" });
    }
    const exists = await pool.query("SELECT id FROM users WHERE username=$1", [username]);
    if (exists.rows.length > 0) {
      return res.json({ success: false, message: "Tên đăng nhập đã tồn tại" });
    }
    const hashed = await bcrypt.hash(password, 10);
    await pool.query(
      "INSERT INTO users(username, password, role, xa_id) VALUES($1,$2,$3,$4)",
      [username, hashed, role, xa_id || null]
    );
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "Lỗi server" });
  }
});

/* Sửa tài khoản (username, role, xa_id) */
app.put("/users/:id", verifyToken("admin"), async (req, res) => {
  try {
    const { username, role, xa_id } = req.body;
    const id = req.params.id;
    const exists = await pool.query("SELECT id FROM users WHERE username=$1 AND id<>$2", [username, id]);
    if (exists.rows.length > 0) {
      return res.json({ success: false, message: "Tên đăng nhập đã tồn tại" });
    }
    await pool.query(
      "UPDATE users SET username=$1, role=$2, xa_id=$3 WHERE id=$4",
      [username, role, xa_id || null, id]
    );
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "Lỗi server" });
  }
});

/* Đổi mật khẩu (admin đổi cho user) */
app.put("/users/:id/password", verifyToken("admin"), async (req, res) => {
  try {
    const { password } = req.body;
    if (!password) {
      return res.json({ success: false, message: "Thiếu mật khẩu" });
    }
    const hashed = await bcrypt.hash(password, 10);
    await pool.query("UPDATE users SET password=$1 WHERE id=$2", [hashed, req.params.id]);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "Lỗi server" });
  }
});

/* Xóa tài khoản */
app.delete("/users/:id", verifyToken("admin"), async (req, res) => {
  try {
    await pool.query("DELETE FROM users WHERE id=$1", [req.params.id]);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "Lỗi server" });
  }
});

app.get('/check-vanban/:xa_id', async (req, res) => {
  const { xa_id } = req.params;

  const ky = await pool.query(
    "SELECT thang, nam FROM ky_baocao WHERE trang_thai='mo' LIMIT 1"
  );

  if (ky.rows.length === 0) {
    return res.json({ exists: false });
  }

  const { thang, nam } = ky.rows[0];

  const result = await pool.query(
    "SELECT * FROM vanban WHERE xa_id=$1 AND thang=$2 AND nam=$3",
    [xa_id, thang, nam]
  );

  res.json({ exists: result.rows.length > 0 });
});

app.get('/xa', async (req, res) => {
  try {
    const result = await pool.query("SELECT id, ten_xa FROM xa ORDER BY ten_xa");
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json([]);
  }
});

app.listen(3000, "0.0.0.0", () => {
  console.log("Server chạy tại http://0.0.0.0:3000");
});