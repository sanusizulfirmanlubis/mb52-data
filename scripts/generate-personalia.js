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

if (!fs.existsSync(folderPath)) {
  throw new Error("Folder personalia tidak ditemukan!");
}

// ======================
// NORMALIZE DATE FUNCTION
// ======================
function normalizeDate(value) {
  if (!value) return "";

  // Already correct format
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  // Excel number date
  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const result = new Date(excelEpoch.getTime() + value * 86400000);
    return result.toISOString().split("T")[0];
  }

  // JS Date object
  if (value instanceof Date) {
    return value.toISOString().split("T")[0];
  }

  return "";
}

fs.readdirSync(folderPath).forEach((file) => {
  if (!file.endsWith(".xlsx")) return;

  console.log("Checking file:", file);

// ======================
// AMBIL NAMA WILAYAH DARI FILE
// ======================
const fileNameWithoutExt = file.replace(".xlsx", "").trim();

// Hapus angka + spasi di depan
const wilayahFromFile = fileNameWithoutExt.replace(/^\d+\s*/, "").trim();
  
  const filePath = path.join(folderPath, file);
  const workbook = XLSX.readFile(filePath);

  if (!workbook.SheetNames.length) {
    throw new Error(`File ${file} tidak memiliki sheet`);
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (!raw.length) {
    console.log(`File ${file} kosong, dilewati...`);
    return;
  }

  const headers = raw[0];

  if (!headers || !Array.isArray(headers)) {
    console.log(`Header tidak ditemukan di file ${file}, dilewati...`);
    return;
  }

  REQUIRED_HEADERS.forEach((req) => {
    if (!headers.includes(req)) {
      throw new Error(`Header "${req}" tidak ditemukan di file ${file}`);
    }
  });

  const rows = XLSX.utils.sheet_to_json(sheet);

  rows.forEach((row, index) => {
    const rowNumber = index + 2;

    const nama = (row["Nama"] || "").toString().trim();

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

      if (!nipp) {
        throw new Error(`Baris ${rowNumber} (${file}): NIPP wajib diisi`);
      }

      if (!/^\d{5}$/.test(nipp)) {
        throw new Error(`Baris ${rowNumber} (${file}): NIPP harus 5 digit angka`);
      }

      if (!pendidikan) {
        throw new Error(`Baris ${rowNumber} (${file}): Pendidikan wajib diisi`);
      }

      if (!mulaiDinas) {
        throw new Error(`Baris ${rowNumber} (${file}): Mulai Dinas wajib diisi`);
      }

      if (!tglLahir) {
        throw new Error(`Baris ${rowNumber} (${file}): Tanggal Lahir wajib diisi`);
      }

      if (!tmtPensiun) {
        throw new Error(`Baris ${rowNumber} (${file}): TMT Pensiun wajib diisi`);
      }
    }
  });

  masterData = masterData.concat(rows);
});

fs.writeFileSync(
  "data/personalia_master.json",
  JSON.stringify(masterData, null, 2)
);

console.log("Master JSON generated successfully");
