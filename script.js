const API_URL = "https://script.google.com/a/macros/essigq.com/s/AKfycbx4YjI_6elYZmYOiW_U2VJrAFzwSvbcoLvy-QeL-EP5t-lIdRY9jdJasL9nHnmtH2n2/exec";

let inventoryData = [];

async function loadInventory() {
  const tableBody = document.getElementById("inventoryTable");

  try {
    const response = await fetch(API_URL);
    inventoryData = await response.json();

    renderTable(inventoryData);
  } catch (error) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="8">Error loading inventory.</td>
      </tr>
    `;
    console.error(error);
  }
}

function renderTable(data) {
  const tableBody = document.getElementById("inventoryTable");

  if (!data || data.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="8">No inventory found.</td>
      </tr>
    `;
    return;
  }

  tableBody.innerHTML = data.map(item => {
    const assignee = item.Assignee || "Available";
    const badgeClass = assignee === "Available" ? "available" : "assigned";

    return `
      <tr>
        <td>${item.Device || ""}</td>
        <td>${item.Brand || ""}</td>
        <td>${item.Model || ""}</td>
        <td>${item.SN || ""}</td>
        <td>${item.Team || ""}</td>
        <td>${item["Internal SN"] || ""}</td>
        <td><span class="badge ${badgeClass}">${assignee}</span></td>
        <td>${item.Labeled || ""}</td>
      </tr>
    `;
  }).join("");
}

document.getElementById("searchInput").addEventListener("input", function () {
  const searchValue = this.value.toLowerCase();

  const filteredData = inventoryData.filter(item => {
    return Object.values(item).some(value =>
      String(value).toLowerCase().includes(searchValue)
    );
  });

  renderTable(filteredData);
});

loadInventory();
