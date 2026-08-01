// =====================================================
// DASHBOARD MODULE
// =====================================================
async function loadTrangThai() {
    const thang = document.getElementById("thang").value;
    const res = await fetch("/dashboard-xa");
    const data = await res.json();
    const tbody = document.querySelector("#bangxa tbody");
}

async function loadLog() {
    const res = await fetch("/log");
    const data = await res.json();
    const tbody = document.querySelector("#logTable tbody");
    tbody.innerHTML = "";
    data.forEach((l, index) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
        <td>${index + 1}</td>
        <td>${l.ten_xa || "Admin"}</td>
        <td>${l.hanh_dong}</td>
        <td>${new Date(l.thoi_gian).toLocaleString("vi-VN")}</td>
    `;
        tbody.appendChild(tr);
    });
}
