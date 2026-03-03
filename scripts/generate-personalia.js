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

const dateRegex = /^\d{4}-\d{2}-\d{2}$/;

fs.readdirSync(folderPath).forEach((file) => {
  if (!file.endsWith(".xlsx")) return;

  console.log("Checking file:", file);

  const filePath = path.join(folderPath, file);
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headers = raw[0];

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
function normalizeDate(value) {
  if (!value) return "";

  // Kalau sudah string YYYY-MM-DD
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  // Kalau Excel date number
  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const result = new Date(excelEpoch.getTime() + value * 86400000);
    return result.toISOString().split("T")[0];
  }

  // Kalau Date object
  if (value instanceof Date) {
    return value.toISOString().split("T")[0];
  }

  return "";
}

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
