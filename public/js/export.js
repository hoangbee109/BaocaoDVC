 // =====================================================
    // EXPORT
    // =====================================================
    async function exportExcel() {
      try {
        const res = await fetch("/export-excel", {
          headers: {
            Authorization: "Bearer " + token
          }
        });
        if (!res.ok) {
          const text = await res.text();
          alert(text);
          return;
        }
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "TongHop.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      } catch (err) {
        console.error(err);
        alert("Không thể xuất Excel");
      }
    }

    async function exportTongDonVi() {
      try {
        const res = await fetch("/export-tong-donvi", {
          headers: {
            Authorization: "Bearer " + token
          }
        });
        if (!res.ok) {
          alert(await res.text());
          return;
        }
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "TongHopDonVi.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      } catch (err) {
        console.error(err);
        alert("Không thể xuất Excel.");
      }
    }

    async function exportDonVi(id) {
      const res = await fetch("/export-donvi-admin?id=" + id, {
        headers: {
          Authorization: "Bearer " + token
        }
      });
      if (!res.ok) {
        const text = await res.text();
        alert(text);
        return;
      }
      const blob = await res.blob();
      const disposition = res.headers.get("Content-Disposition");
      let fileName = "BaoCao.xlsx";
      if (disposition) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match) {
          fileName = match[1];
        }
      }
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    }
