   // =====================================================
    // PROGRESS MODULE
    // =====================================================
    async function loadProgress() {
      const res = await fetch("/dashboard-xa");
      const data = await res.json();
      const tong = data.length;
      const daGui = data.filter(x => x.trang_thai === "Đã gửi").length;
      const chuaGui = data.filter(x => x.trang_thai !== "Đã gửi").length;
      document.getElementById("tongDV").innerText = tong;
      document.getElementById("daGui").innerText = daGui;
      document.getElementById("chuaGui").innerText = chuaGui;
      document.getElementById("tiLeGui").innerText = Math.round(daGui * 100 / tong) + "%";
      const tbody = document.querySelector("#tableProgress tbody");
      tbody.innerHTML = "";
      data.forEach((x, index) => {
        let badge = ""; if (x.trang_thai === STATUS_SUBMITTED) {
          badge = `<span class="badge-ok">🟢 Đã gửi</span>`;
        } else if (x.trang_thai === STATUS_EDITING) {
          badge = `<span class="badge-warning">🟡 Đang nhập</span>`;
        } else { badge = `<span class="badge-no">🔴 Chưa nhập</span>`; }
        let action = "";
        if (x.trang_thai === STATUS_SUBMITTED) {
          action = `
    <button class="btn-open" onclick="moBaoCao(${x.id})">
      Mở lại báo cáo
    </button>
    <button class="btn-excel"
      onclick="exportDonVi(${x.id})">
      Excel
    </button>
  `;
        } else if (x.trang_thai === STATUS_EDITING) {
          action = `
    <button class="btn-excel"
      onclick="exportDonVi(${x.id})">
      Excel
    </button>
  `;
        } else {
          action = `
    <button class="btn-excel-disable" disabled>
      Excel
    </button>
  `;
        }
        tbody.innerHTML += `
<tr>
<td>${index + 1}</td>
<td>${x.ten_xa}</td>
<td><div class="action-group">${action}</div></td>
<td style="text-align:center">
${badge}
</td>
<td style="text-align:center">
${x.cap_nhat_cuoi
            ? new Date(x.cap_nhat_cuoi)
              .toLocaleString("vi-VN", {
                hour: "2-digit",
                minute: "2-digit",
                day: "2-digit",
                month: "2-digit",
                year: "numeric"
              })
              .replace(",", " ngày")
            : "-"
          }
</td>
</tr>`;
      });
    }

    function filterProgress() {
      let visible = 0;
      const keyword = document.getElementById("searchProgress").value.trim().toLowerCase();
      const status = document.getElementById("filterProgress").value.trim();
      const rows = document.querySelectorAll("#tableProgress tbody tr");
      rows.forEach(row => {
        const unitName = row.cells[COL_UNIT].innerText.toLowerCase();
        const statusText = row.cells[COL_STATUS].innerText;
        const okTen = unitName.includes(keyword);
        const okStatus = status === STATUS_ALL || statusText.includes(status);
        if (okTen && okStatus) {
          visible++;
        }
        row.style.display = okTen && okStatus ? "" : "none";
      });
      const totalRows = rows.length;
      document.getElementById("progressCount").innerText = ` ${visible} /  ${totalRows} đơn vị `;
    }

    function copyChuaNhap() {
      const rows = [...document.querySelectorAll("#tableProgress tbody tr")];
      let text = "";
      rows.forEach(r => {
        if (r.innerText.includes("Chưa nhập")) {
          text += r.children[1].innerText + "\n";
        }
      });
      navigator.clipboard.writeText(text);
      alert("Đã copy danh sách.");
    }
