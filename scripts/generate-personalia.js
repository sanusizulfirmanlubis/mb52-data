const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const REQUIRED_HEADERS = [
  "Wilayah",
  "Nama",
  "NIPP",
  "Level Jabatan",
  "Tingkat Jabatan",
  "Posisi",
  "UPT",
  "TMT Posisi (YYYY-MM-DD)",
  "Grade",
  "Pendidikan",
  "Mulai Dinas (YYYY-MM-DD)",
  "Tanggal Lahir (YYYY-MM-DD)",
  "TMT Pensiun (YYYY-MM-DD)",
  "No PRP",
  "Berlaku PRP (YYYY-MM-DD)",
  "Tingkat PRP",
  "No PMP",
  "Berlaku PMP (YYYY-MM-DD)",
  "Tingkat PMP",
  "No WA"
];

const folderPath = "data/personalia";

let masterData = [];
let errorRows = [];
let stats = {};

if (!fs.existsSync(folderPath)) {
  throw new Error("Folder personalia tidak ditemukan!");
}

// ======================
// DATE NORMALIZER
// ======================
function normalizeDate(value) {

  if (!value) return "";

  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const result = new Date(excelEpoch.getTime() + value * 86400000);
    return result.toISOString().split("T")[0];
  }

  if (value instanceof Date) {
    return value.toISOString().split("T")[0];
  }

  return "";
}

// ======================
// GET FILE LIST
// ======================
const files = fs
  .readdirSync(folderPath)
  .filter((file) => file.endsWith(".xlsx"));

console.log(`\nProcessing ${files.length} wilayah...\n`);

// ======================
// PROCESS FILE
// ======================
files.forEach((file) => {

  console.log(`📂 Checking file: ${file}`);

  const wilayahFromFile = file.replace(".xlsx", "").trim();

  const filePath = path.join(folderPath, file);

  const workbook = XLSX.readFile(filePath, {
    cellFormula: false,
    cellHTML: false
  });

  if (!workbook.SheetNames.length) {
    console.log(`⚠ File ${file} tidak memiliki sheet`);
    return;
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  // FAST PARSE (hanya sekali)
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true
  });

  if (!rows.length) {
    console.log(`⚠ File ${file} kosong`);
    return;
  }

  const headers = Object.keys(rows[0]);

  REQUIRED_HEADERS.forEach((req) => {

    if (!headers.includes(req)) {

      errorRows.push({
        file,
        row: "HEADER",
        message: `Header "${req}" tidak ditemukan`
      });

    }

  });

  let validCount = 0;

  rows.forEach((row, index) => {

    const rowNumber = index + 2;

    try {

      const nama = (row["Nama"] || "").toString().trim();

      const wilayah = (row["Wilayah"] || "").toString().trim();

      if (wilayah !== wilayahFromFile) {

        errorRows.push({
          file,
          row: rowNumber,
          message: `Wilayah "${wilayah}" tidak sesuai dengan nama file`
        });

        return;

      }

      let nippRaw = row["NIPP"];

      if (typeof nippRaw === "number") {
        nippRaw = Math.floor(nippRaw).toString();
      }

      const nipp = (nippRaw || "").toString().trim();

      const pendidikan = (row["Pendidikan"] || "").toString().trim();

      const mulaiDinas = normalizeDate(row["Mulai Dinas (YYYY-MM-DD)"]);
      const tglLahir = normalizeDate(row["Tanggal Lahir (YYYY-MM-DD)"]);
      const tmtPensiun = normalizeDate(row["TMT Pensiun (YYYY-MM-DD)"]);

      const isVacant = nama.toLowerCase() === "vacant";

      if (!isVacant) {

        if (!nipp) throw "NIPP wajib diisi";

        if (!/^\d{5}$/.test(nipp)) throw "NIPP harus 5 digit angka";

        if (!pendidikan) throw "Pendidikan wajib diisi";

        if (!mulaiDinas) throw "Mulai Dinas wajib diisi";

        if (!tglLahir) throw "Tanggal Lahir wajib diisi";

        if (!tmtPensiun) throw "TMT Pensiun wajib diisi";

      }

      masterData.push(row);

      validCount++;

    } catch (err) {

      errorRows.push({
        file,
        row: rowNumber,
        message: err.toString()
      });

    }

  });

  stats[file] = validCount;

});

// ======================
// WRITE MASTER DATA
// ======================
fs.writeFileSync(
  "data/personalia_master.json",
  JSON.stringify(masterData)
);

// ======================
// WRITE ERROR REPORT
// ======================
if (errorRows.length > 0) {

  fs.writeFileSync(
    "data/personalia_error_report.json",
    JSON.stringify(errorRows, null, 2)
  );

}

// ======================
// LOG SUMMARY
// ======================
console.log("\n================ SUMMARY ================");

console.log(`Total Data Valid : ${masterData.length}`);

console.log("\nData per wilayah:");

Object.keys(stats).forEach((file) => {
  console.log(`${file} : ${stats[file]} pegawai`);
});

if (errorRows.length > 0) {

  console.log("\n⚠ ERROR REPORT");

  errorRows.forEach((err) => {
    console.log(`File: ${err.file} | Baris: ${err.row} | ${err.message}`);
  });

  console.log(`\nTotal Error: ${errorRows.length}`);

}

console.log("=========================================");

console.log("\n✅ Master JSON generated successfully\n");
