let workbookData = null;
let convertedData = null;

document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    workbookData = jsonData;
    document.getElementById('convertBtn').disabled = false;

    renderTable(jsonData); // عرض الملف الأصلي
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById('convertBtn').addEventListener('click', function () {
  if (!workbookData) return;

  // ✅ البيانات الأساسية من المواقع الثابتة
  let studentName = workbookData[4]?.[1] || ""; // B5
  let subject = workbookData[6]?.[1] || "";     // B7
  let goal = workbookData[8]?.[1] || "";        // B9

  // ✅ التاريخ (Regex)
  let date = "";
  workbookData.forEach(row => {
    row.forEach(cell => {
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(cell)) {
        date = cell;
      }
    });
  });

  // ✅ البنود تبدأ من الصف 13
  let startIndex = 12;

  convertedData = [["الطالب", "التاريخ", "المادة", "الهدف", "بنود التقييم/التقييم", "بنود التقييم/درجة التقييم", "بنود التقييم/ملاحظات التقييم"]];

  for (let i = startIndex; i < workbookData.length; i++) {
    let row = workbookData[i];

    let evaluationItem = row[0] || ""; // العمود A
    let grade = row[4] || row[5] || ""; // الدرجة E أو F
    let notes = row[6] || row[7] || ""; // الملاحظات G أو H

    // وقف عند أول صف فاضي تمامًا
    if (!evaluationItem && !grade && !notes) break;

    // أضف الصف
    convertedData.push([
      i === startIndex ? studentName : "",
      i === startIndex ? date : "",
      i === startIndex ? subject : "",
      i === startIndex ? goal : "",
      evaluationItem,
      grade,
      notes
    ]);

    // ✅ إذا البند يحتوي على "ملخص" توقف فورًا
    if (String(evaluationItem).trim().includes("ملخص")) {
      break;
    }
  }

  // ✅ عرض النتيجة
  renderTable(convertedData);

  document.getElementById('downloadBtn').disabled = false;
  alert("✅ Conversion done! Stops after 'ملخص'.");
});

document.getElementById('downloadBtn').addEventListener('click', function () {
  if (!convertedData) return;
  const ws = XLSX.utils.aoa_to_sheet(convertedData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Converted Report");
  XLSX.writeFile(wb, "Converted_Report.xlsx");
});

function renderTable(data) {
  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  if (data.length > 0) {
    // إضافة رؤوس الأعمدة
    data[0].forEach(header => {
      const th = document.createElement('th');
      th.textContent = header || '';
      tableHead.appendChild(th);
    });

    // إضافة الصفوف
    data.slice(1).forEach(row => {
      const tr = document.createElement('tr');
      row.forEach(cell => {
        const td = document.createElement('td');
        td.textContent = cell || '';
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });
  }
}






// let workbookData = null;
// let convertedData = null;

// document.getElementById('excelFile').addEventListener('change', function (e) {
//   const file = e.target.files[0];
//   if (!file) return;

//   const reader = new FileReader();
//   reader.onload = function (event) {
//     const data = new Uint8Array(event.target.result);
//     const workbook = XLSX.read(data, { type: 'array' });

//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

//     workbookData = jsonData;
//     document.getElementById('convertBtn').disabled = false;

//     renderTable(jsonData); // عرض الملف الأصلي
//   };
//   reader.readAsArrayBuffer(file);
// });

// document.getElementById('convertBtn').addEventListener('click', function () {
//   if (!workbookData) return;

//   let studentName = "";
//   let date = "";
//   let subject = "";
//   let goal = "";

//   // ✅ التعرف على البيانات الأساسية بالبحث عن الكلمات
//   workbookData.forEach(row => {
//     row.forEach((cell, index) => {
//       if (typeof cell === "string" && cell.includes("الطالب")) {
//         studentName = row[index + 1] || cell.replace("الطالب:", "").trim();
//       }
//       if (/^\d{2}\/\d{2}\/\d{4}$/.test(cell)) {
//         date = cell;
//       }
//       if (typeof cell === "string" && cell.trim() === "المادة") {
//         subject = row[index - 1] || "";
//       }
//       if (typeof cell === "string" && cell.includes("هدف")) {
//         goal = row[index + 1] || row[index] || "";
//       }
//     });
//   });

//   // ✅ البنود تبدأ من الصف 13 (Index 12)
//   let startIndex = 12;

//   convertedData = [["الطالب", "التاريخ", "المادة", "الهدف", "بنود التقييم/التقييم", "بنود التقييم/درجة التقييم", "بنود التقييم/ملاحظات التقييم"]];

//   for (let i = startIndex; i < workbookData.length; i++) {
//     let row = workbookData[i];
//     let evaluationItem = row[0] || ""; // العمود A
//     let grade = row[4] || "";          // العمود E
//     let notes = row[6] || "";          // العمود G (تعديل جديد)

//     // وقف عند أول صف فاضي تمامًا
//     if (!evaluationItem && !grade && !notes) break;

//     convertedData.push([
//       i === startIndex ? studentName : "",
//       i === startIndex ? date : "",
//       i === startIndex ? subject : "",
//       i === startIndex ? goal : "",
//       evaluationItem,
//       grade,
//       notes
//     ]);
//   }

//   // ✅ عرض النتيجة
//   renderTable(convertedData);

//   document.getElementById('downloadBtn').disabled = false;
//   alert("✅ Conversion done! Notes now read from column G.");
// });

// document.getElementById('downloadBtn').addEventListener('click', function () {
//   if (!convertedData) return;
//   const ws = XLSX.utils.aoa_to_sheet(convertedData);
//   const wb = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(wb, ws, "Converted Report");
//   XLSX.writeFile(wb, "Converted_Report.xlsx");
// });

// function renderTable(data) {
//   const tableHead = document.getElementById('tableHead');
//   const tableBody = document.getElementById('tableBody');
//   tableHead.innerHTML = '';
//   tableBody.innerHTML = '';

//   if (data.length > 0) {
//     // إضافة رؤوس الأعمدة
//     data[0].forEach(header => {
//       const th = document.createElement('th');
//       th.textContent = header || '';
//       tableHead.appendChild(th);
//     });

//     // إضافة الصفوف
//     data.slice(1).forEach(row => {
//       const tr = document.createElement('tr');
//       row.forEach(cell => {
//         const td = document.createElement('td');
//         td.textContent = cell || '';
//         tr.appendChild(td);
//       });
//       tableBody.appendChild(tr);
//     });
//   }
// }
