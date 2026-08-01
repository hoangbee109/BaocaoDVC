// =====================================================
// PERMISSION MODULE (TTHC)
// =====================================================
let dsNhom = [];
let dsChitieu = [];
let dsPhanQuyen = [];
async function loadPhanQuyen() {
    try {
        const [rNhom, rCT, rPQ] = await Promise.all([
            fetch("/nhom-tthc", { headers }),
            fetch("/admin/chitieu", { headers }),
            fetch("/phanquyen-tthc", { headers })
        ]);
        const nhom = await rNhom.json();
        const chitieu = await rCT.json();
        const pq = await rPQ.json();
        if (!nhom.success || !chitieu.success || !pq.success) {
            alert("Không tải được dữ liệu phân quyền.");
            return;
        }
        dsNhom = nhom.data;
        dsChitieu = chitieu.data;
        dsPhanQuyen = pq.data;
        buildBangPhanQuyen();
    } catch (err) {
        console.error(err);
        alert(err.message);
    }
}
function buildBangPhanQuyen() {
    const table = document.getElementById("tblPhanQuyen");
    table.innerHTML = "";
    // Header
    const thead = document.createElement("thead");
    let html = `
        <tr>
            <th style="width:60px">STT</th>
            <th style="min-width:420px">Tên thủ tục hành chính</th>
    `;
    dsNhom.forEach(n => {
        html += `
            <th>${n.ma_nhom}</th>
        `;
    });
    html += "</tr>";
    thead.innerHTML = html;
    table.appendChild(thead);
    // Body
    const tbody = document.createElement("tbody");
    dsChitieu.forEach(ct => {
        let tr = document.createElement("tr");
        let row = `<td>${ct.stt}</td> <td>${ct.ten_chitieu}</td>`;
        dsNhom.forEach(n => {
            const checked = dsPhanQuyen.some(p =>
                p.ma_nhom == n.ma_nhom &&
                p.chitieu_id == ct.id
            );
            row += `
                <td style="text-align:center">
                    <input
                        type="checkbox"
                        class="chkPQ"
                        data-nhom="${n.ma_nhom}"
                        data-id="${ct.id}"
                        ${checked ? "checked" : ""}>
                </td>
            `;
        });
        tr.innerHTML = row;
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
}
async function savePhanQuyen() {
    const ds = [];
    document.querySelectorAll(".chkPQ").forEach(chk => {
        if (chk.checked) {
            ds.push({
                ma_nhom: chk.dataset.nhom,
                chitieu_id: Number(chk.dataset.id)
            });
        }
    });
    const res = await fetch("/save-phanquyen", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token
        },
        body: JSON.stringify({
            data: ds
        })
    });
    const json = await res.json();
    if (json.success) {
        alert("✅ Đã lưu phân quyền!");
    } else {
        alert(json.message);
    }
}
// =====================================================
// EVENT LISTENERS
// =====================================================
document.getElementById("txtSearchTTHC")
    ?.addEventListener("input", function () {

        const key = this.value.toLowerCase();

        document
            .querySelectorAll("#tblPhanQuyen tbody tr")
            .forEach(tr => {

                const ten = tr.cells[1]
                    .innerText
                    .toLowerCase();

                tr.style.display =
                    ten.includes(key)
                        ? ""
                        : "none";

            });

    });