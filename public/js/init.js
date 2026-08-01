async function loadAll() {
    await loadProgress();
}
window.addEventListener("DOMContentLoaded", async () => {
    await loadAll();
    loadLog();
});
function showTab(tab) {
    document.querySelectorAll(".portal-page")
        .forEach(page => page.style.display = "none");
    const map = {
        dashboard: "tabDashboard",
        ky: "tabKy",
        phanquyen: "tabPhanquyen",
        progress: "tabProgress",
        users: "tabUsers",
        logs: "tabLogs",
        documents: "tabDocuments"
    };
    const page = document.getElementById(map[tab]);
    if (page) {
        page.style.display = "block";
    }
    if (tab === "ky") {
        loadKyBaoCao();
    }
    if (tab === "phanquyen") {
        loadPhanQuyen();
    }
    if (tab === "progress") {
        loadProgress();
    }
    if (tab === "users") {
        loadUsers();
    }
    if (tab === "logs") {
        loadLog();
    }
    document.querySelectorAll(".sidebar a")
        .forEach(a => a.classList.remove("active"));
    if (event && event.target) {
        event.target.classList.add("active");
    }
}