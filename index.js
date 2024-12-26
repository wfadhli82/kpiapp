document.addEventListener("DOMContentLoaded", () => {
  const measurementDropdown = document.getElementById("measurement");
  const timeFields = document.getElementById("timeFields");
  const quantityFields = document.getElementById("quantityFields");
  const percentageFields = document.getElementById("percentageFields");
  const kpiTableBody = document.getElementById("kpiTable").querySelector("tbody");
  const downloadExcelButton = document.getElementById("downloadExcel");
  const uploadExcelButton = document.getElementById("uploadExcel");
  const fileInput = document.getElementById("fileInput");
  const kpiData = [];
  let editIndex = null;

  const measurementLabels = {
    time: "Masa",
    quantity: "Kuantiti",
    percentage: "Peratus",
  };

  measurementDropdown.addEventListener("change", () => {
    const selectedValue = measurementDropdown.value;

    timeFields.classList.add("hidden");
    quantityFields.classList.add("hidden");
    percentageFields.classList.add("hidden");

    if (selectedValue === "time") {
      timeFields.classList.remove("hidden");
    } else if (selectedValue === "quantity") {
      quantityFields.classList.remove("hidden");
    } else if (selectedValue === "percentage") {
      percentageFields.classList.remove("hidden");
    }
  });

  uploadExcelButton.addEventListener("click", () => {
    fileInput.click();
  });

  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      jsonData.forEach((row) => {
        const rowData = {
          department: row.Bahagian || "",
          activity: row.Aktiviti || "",
          percentageAchievement: row["Peratus Pencapaian"] || "",
          measurement: row.Ukuran || "",
          target: row.Sasaran || "",
          achieved: row.Capai || "",
          totalValue: row["Jumlah Keseluruhan"] || "",
          achievedValue: row["Jumlah Dicapai"] || "",
        };
        kpiData.push(rowData);
        addRowToTable(rowData);
      });
    };
    reader.readAsArrayBuffer(file);
  });

  downloadExcelButton.addEventListener("click", () => {
    if (kpiData.length === 0) {
      alert("Tiada data untuk dimuat turun!");
      return;
    }

    const headers = [
      { header: "Bahagian", key: "department" },
      { header: "Aktiviti", key: "activity" },
      { header: "Ukuran", key: "measurement" },
      { header: "Sasaran", key: "target" },
      { header: "Capai", key: "achieved" },
      { header: "Peratus Pencapaian", key: "percentageAchievement" },
      { header: "Jumlah Keseluruhan", key: "totalValue" },
      { header: "Jumlah Dicapai", key: "achievedValue" },
    ];

    const structuredData = kpiData.map((row) => {
      return headers.reduce((acc, header) => {
        acc[header.header] = row[header.key] || "";
        return acc;
      }, {});
    });

    const worksheet = XLSX.utils.json_to_sheet(structuredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "KPI Data");

    XLSX.writeFile(workbook, "KPI_Data.xlsx");
  });

  function addRowToTable(rowData) {
    const newRow = document.createElement("tr");
    newRow.innerHTML = `
      <td>${kpiTableBody.rows.length + 1}</td>
      <td>${rowData.department}</td>
      <td>${rowData.activity}</td>
      <td>${rowData.percentageAchievement}</td>
      <td>${rowData.target}</td>
      <td>${rowData.achieved}</td>
      <td>${rowData.totalValue}</td>
      <td>${rowData.achievedValue}</td>
      <td>${rowData.measurement}</td>
      <td><button class="edit-btn"><i class="fas fa-edit"></i></button></td>
    `;
    kpiTableBody.appendChild(newRow);

    const editButton = newRow.querySelector(".edit-btn");
    editButton.addEventListener("click", () => {
      const rowIndex = newRow.rowIndex - 1; // Row index tanpa header
      populateFormForEdit(kpiData[rowIndex], rowIndex);
    });
  }

  function populateFormForEdit(rowData, index) {
    document.getElementById("department").value = rowData.department;
    document.getElementById("activity").value = rowData.activity;
    document.getElementById("measurement").value = Object.keys(measurementLabels).find(key => measurementLabels[key] === rowData.measurement);

    if (rowData.measurement === "Masa") {
      document.getElementById("targetDate").value = rowData.target;
      document.getElementById("achievedDate").value = rowData.achieved;
    } else if (rowData.measurement === "Kuantiti") {
      document.getElementById("targetQuantity").value = rowData.target;
      document.getElementById("achievedQuantity").value = rowData.achieved;
    } else if (rowData.measurement === "Peratus") {
      document.getElementById("targetPercentage").value = parseFloat(rowData.target.replace("%", "")) || "";
      document.getElementById("targetValue").value = rowData.totalValue;
      document.getElementById("achievedValue").value = rowData.achievedValue;
    }

    kpiForm.dataset.editIndex = index;
  }

  kpiForm.addEventListener("submit", (e) => {
    e.preventDefault();

    const formData = new FormData(kpiForm);
    const data = Object.fromEntries(formData.entries());

    let target = "N/A";
    let achieved = "N/A";
    let percentageAchievement = "N/A";
    let totalValue = "Tidak Berkaitan";
    let achievedValue = "Tidak Berkaitan";

    if (data.measurement === "time") {
      const targetDate = new Date(data.targetDate);
      const achievedDate = new Date(data.achievedDate);

      if (!isNaN(targetDate) && !isNaN(achievedDate)) {
        const daysLate = Math.max(0, (achievedDate - targetDate) / (1000 * 60 * 60 * 24));
        const totalDaysInYear = targetDate.getFullYear() % 4 === 0 ? 366 : 365;

        percentageAchievement =
          daysLate === 0
            ? "100%"
            : Math.max(0, (100 - (daysLate / totalDaysInYear) * 100)).toFixed(2) + "%";

        target = data.targetDate;
        achieved = data.achievedDate;
      }
    } else if (data.measurement === "quantity") {
      target = parseFloat(data.targetQuantity) || 0;
      achieved = parseFloat(data.achievedQuantity) || 0;

      if (achieved >= target) {
        percentageAchievement = "100%";
      } else {
        percentageAchievement =
          target > 0 ? ((achieved / target) * 100).toFixed(2) + "%" : "0%";
      }
    } else if (data.measurement === "percentage") {
      const targetValue = parseFloat(data.targetValue) || 0;
      const achievedValueRaw = parseFloat(data.achievedValue) || 0;
      const targetPercentage = parseFloat(data.targetPercentage) || 0;

      if (targetValue === 0) {
        percentageAchievement = "Tiada Pengiraan";
        totalValue = "Tidak Berkaitan";
        achievedValue = "Tidak Berkaitan";
      } else {
        totalValue = targetValue.toFixed(2);
        achievedValue = achievedValueRaw.toFixed(2);

        const calculatedPercentage =
          targetValue > 0 ? (achievedValueRaw / targetValue) * 100 : 0;

        percentageAchievement =
          calculatedPercentage >= targetPercentage
            ? "100%"
            : ((calculatedPercentage / targetPercentage) * 100).toFixed(2) + "%";

        target = targetPercentage + "%";
        achieved = calculatedPercentage.toFixed(2) + "%";
      }
    }

    const rowData = {
      department: data.department,
      activity: data.activity,
      percentageAchievement,
      measurement: measurementLabels[data.measurement],
      target,
      achieved,
      totalValue,
      achievedValue,
    };

    if (kpiForm.dataset.editIndex !== undefined) {
      const index = parseInt(kpiForm.dataset.editIndex, 10);
      kpiData[index] = rowData;

      const row = kpiTableBody.rows[index];
      row.cells[1].textContent = rowData.department;
      row.cells[2].textContent = rowData.activity;
      row.cells[3].textContent = rowData.percentageAchievement;
      row.cells[4].textContent = rowData.target;
      row.cells[5].textContent = rowData.achieved;
      row.cells[6].textContent = rowData.totalValue;
      row.cells[7].textContent = rowData.achievedValue;
      row.cells[8].textContent = rowData.measurement;

      delete kpiForm.dataset.editIndex;
    } else {
      kpiData.push(rowData);
      addRowToTable(rowData);
    }

    kpiForm.reset();
  });
});
