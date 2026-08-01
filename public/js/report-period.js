// =====================================================
// REPORT PERIOD MODULE
// =====================================================
async function loadKyBaoCao() {
    const res = await fetch("/ky-baocao-list", {
        headers
    });
    const json = await res.json();
    if (!json.success) {
        alert(json.message);
        return;
    }
    renderKyBaoCao(json.data);
}

function renderKyBaoCao(list) {
    if (!list || list.length === 0) return;
    // KỲ ĐANG MỞ
    const current = list.find(x => x.trang_thai === "mo");
    if (current) {
        document.querySelector(".ky-month").innerHTML =
            "Tháng " + current.thang + " / " + current.nam;
        document.querySelector(".ky-date").innerHTML =
            formatDate(current.ngay_bat_dau)
            + " → "
            + formatDate(current.ngay_ket_thuc);
    }
    // DANH SÁCH
    const tbody = document.getElementById("kyTableBody");
    tbody.innerHTML = "";
    list.forEach((k, index) => {
        tbody.innerHTML += `
<tr>
<td>${index + 1}</td>
<td>${k.thang}</td>
<td>${k.nam}</td>
<td>${formatDate(k.ngay_bat_dau)}</td>
<td>${formatDate(k.ngay_ket_thuc)}</td>
<td>
${k.trang_thai == "mo"
                ?
                '🟢 Đang mở'
                :
                '🔴 Đã đóng'}
</td>
<td>
<button>
Chi tiết
</button>
</td>
</tr>
`;
    });
    document.getElementById("newMonth").value = current.thang + 1;
    document.getElementById("newYear").value = current.nam;
    if (current.thang == 12) {
        document.getElementById("newMonth").value = 1;
        document.getElementById("newYear").value =
            current.nam + 1;
    }
    updateKyPreview();
}

function updateKyPreview() {
    const month = Number(document.getElementById("newMonth").value);
    const year = Number(document.getElementById("newYear").value);
    let startMonth = month - 1;
    let startYear = year;
    if (startMonth === 0) {
        startMonth = 12;
        startYear--;
    }
    const start = `${startYear}-${String(startMonth).padStart(2, "0")}-13`;
    const end = `${year}-${String(month).padStart(2, "0")}-12`;
    document.getElementById("startDate").value = start;
    document.getElementById("endDate").value = end;
}

async function chuyenKyBaoCao() {
    if (!confirm("Bạn có chắc chắn muốn chuyển sang kỳ báo cáo mới?")) {
        return;
    }
    const body = {
        thang: Number(document.getElementById("newMonth").value),
        nam: Number(document.getElementById("newYear").value),
        ngay_bat_dau: document.getElementById("startDate").value,
        ngay_ket_thuc: document.getElementById("endDate").value
    };
    const res = await fetch("/chuyen-ky-baocao", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token
        },
        body: JSON.stringify(body)
    });
    const json = await res.json();
    if (!json.success) {
        alert(json.message);
        return;
    }
    alert(json.message);
    await loadKyBaoCao();
}

async function moBaoCao(xa_id) {
    if (!confirm("Cho phép đơn vị chỉnh sửa lại báo cáo?"))
        return;
    const res = await fetch("/mo-lai-baocao", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token
        },
        body: JSON.stringify({
            xa_id
        })
    });
    const json = await res.json();
    if (json.success) {
        alert("Đã mở lại báo cáo.");
        loadProgress();
        loadLog();
    } else {
        alert(json.message);
    }
}

// =====================================================
// EVENT LISTENERS
// =====================================================

document.getElementById("newMonth")
    ?.addEventListener("change", updateKyPreview);

document.getElementById("newYear")
    ?.addEventListener("input", updateKyPreview);

document.getElementById("btnCreateKy")
    ?.addEventListener("click", chuyenKyBaoCao);