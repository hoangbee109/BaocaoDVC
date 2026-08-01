// =====================================================
// USERS MODULE
// =====================================================
let allUsers = [];
let allXa = [];
async function loadUsers() {
    const [resUsers, resXa] = await Promise.all([
        fetch("/users"),
        fetch("/danh-sach-xa")
    ]);
    const users = await resUsers.json();
    allUsers = users;
    allXa = await resXa.json();
    renderUsers();
    populateXaDropdown();
}
function populateXaDropdown() {
    const sel = document.getElementById("inputXaId");
    sel.innerHTML = '<option value="">-- Không chọn --</option>';
    allXa.forEach(x => {
        sel.innerHTML += '<option value="' + x.id + '">' + x.ten_xa + '</option>';
    });
}
function renderUsers() {
    const keyword = (document.getElementById("searchUser").value || "").toLowerCase();
    const tbody = document.querySelector("#tableUsers tbody");
    tbody.innerHTML = "";
    const filtered = allUsers.filter(u =>
        u.username.toLowerCase().includes(keyword) ||
        (u.ten_xa || "").toLowerCase().includes(keyword)
    );
    filtered.forEach(u => {
        const roleName = u.role === "admin" ? "Admin" : u.role === "viewer" ? "Chỉ xem" : "Đơn vị";
        const tr = document.createElement("tr");
        tr.innerHTML = `
<td>${u.id}</td>
<td>${u.username}</td>
<td>${roleName}</td>
<td>${u.ten_xa || "-"}</td>
<td>
<button class="btn-warning btn-sm" onclick="openEditUser(${u.id})">Sửa</button>
<button class="btn-warning btn-sm" onclick="openChangePass(${u.id},'${u.username}')">Reset mật khẩu</button>
<button class="btn-danger btn-sm" onclick="deleteUser(${u.id},'${u.username}')">Xoá</button>
</td>
`;
        tbody.appendChild(tr);
    });
}
function openAddUser() {
    document.getElementById("modalUserTitle").innerText = "Thêm tài khoản";
    document.getElementById("editUserId").value = "";
    document.getElementById("inputUsername").value = "";
    document.getElementById("inputPassword").value = "";
    document.getElementById("inputRole").value = "xa";
    document.getElementById("inputXaId").value = "";
    document.getElementById("passwordFields").style.display = "block";
    document.getElementById("modalUser").style.display = "flex";
}
function openEditUser(id) {
    const u = allUsers.find(x => x.id === id);
    if (!u) return;
    document.getElementById("modalUserTitle").innerText = "Sửa tài khoản: " + u.username;
    document.getElementById("editUserId").value = u.id;
    document.getElementById("inputUsername").value = u.username;
    document.getElementById("inputRole").value = u.role;
    populateXaDropdown();
    document.getElementById("inputXaId").value = u.xa_id || "";
    document.getElementById("passwordFields").style.display = "none";
    document.getElementById("modalUser").style.display = "flex";
}
async function saveUser() {
    const id = document.getElementById("editUserId").value;
    const username = document.getElementById("inputUsername").value.trim();
    const role = document.getElementById("inputRole").value;
    const xa_id = document.getElementById("inputXaId").value;
    if (!username) {
        alert("Nhập tên đăng nhập");
        return;
    }
    if (id) {
        /* Sửa */
        const res = await fetch("/users/" + id, {
            method: "PUT",
            headers: {
                "Content-Type": "application/json",
                Authorization: "Bearer " + token
            },
            body: JSON.stringify({ username, role, xa_id: xa_id || null })
        });
        const data = await res.json();
        if (!data.success) {
            alert(data.message || "Lỗi");
            return;
        }
        alert("Đã cập nhật tài khoản");
    } else {
        /* Thêm */
        const password = document.getElementById("inputPassword").value.trim();
        if (!password) {
            alert("Nhập mật khẩu");
            return;
        }
        const res = await fetch("/users", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: "Bearer " + token
            },
            body: JSON.stringify({ username, password, role, xa_id: xa_id || null })
        });
        const data = await res.json();
        if (!data.success) {
            alert(data.message || "Lỗi");
            return;
        }
        alert("Đã thêm tài khoản");
    }
    closeModal("modalUser");
    loadUsers();
    document.getElementById("searchUser").value = "";
}
async function deleteUser(id, username) {
    if (!confirm("Xoá tài khoản '" + username + "'?")) return;
    const res = await fetch("/users/" + id, { method: "DELETE" });
    const data = await res.json();
    if (!data.success) {
        alert(data.message || "Lỗi");
        return;
    }
    alert("Đã xoá tài khoản thành công.");
    loadUsers();
}
function openChangePass(id, username) {
    document.getElementById("passUserId").value = id;
    document.getElementById("passUserName").innerText = "Tài khoản: " + username;
    document.getElementById("inputNewPassword").value = "";
    document.getElementById("modalPassword").style.display = "flex";
}
async function savePassword() {
    const id = document.getElementById("passUserId").value;
    const password = document.getElementById("inputNewPassword").value.trim();
    if (!password) {
        alert("Nhập mật khẩu mới");
        return;
    }
    const res = await fetch("/users/" + id + "/password", {
        method: "PUT",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token
        },
        body: JSON.stringify({ password })
    });
    const data = await res.json();
    if (!data.success) {
        alert(data.message || "Lỗi");
        return;
    }
    alert("Đã đổi mật khẩu");
    closeModal("modalPassword");
}
async function resetPass(xa_id) {
    if (!confirm("Reset mật khẩu về 123456?")) return;
    await fetch("/reset-mat-khau", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token
        },
        body: JSON.stringify({ xa_id })
    });
    alert("Đã reset mật khẩu");
}
function closeModal(id) {
    document.getElementById(id).style.display = "none";
}
