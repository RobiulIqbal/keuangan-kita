/**
 * ============================================================================
 * KEUANGAN KITA • GOOGLE SHEETS TWO-WAY SYNC API (Next.js / Vercel Serverless)
 * ============================================================================
 * Endpoint: /api/sheets?period=2026-04&key=YOUR_HOUSEHOLD_SECRET_KEY
 *
 * Fitur Utama:
 * 1. Zero-Login Service Account Auth (JWT via googleapis) + URL Secret Key Guard.
 * 2. Anchor-Based Parser untuk membaca blok tabel dinamis pada sheet D-[BULAN]
 *    (Sumber Pendapatan, Fixed Cost, Subscription & Debt, Nabung/Saving).
 * 3. Dual Sub-Ledger Parser & Safe Row Appender untuk sheet T-[BULAN]
 *    (Tabel Pengeluaran Sehari-Hari & Tabel Jajanan/Nongkrong) tanpa menimpa
 *    rumus SUM agregasi.
 * 4. Auto-Provisioning tab D-[BULAN] & T-[BULAN] jika bulan baru belum ada.
 * ============================================================================
 */

import { google } from "googleapis";

const MONTH_CODES = {
  "01": "JAN",
  "02": "FEB",
  "03": "MAR",
  "04": "APR",
  "05": "MAY",
  "06": "JUN",
  "07": "JUL",
  "08": "AUG",
  "09": "SEP",
  10: "OCT",
  11: "NOV",
  12: "DEC",
};

/**
 * Inisialisasi Google Sheets Client menggunakan Service Account (.env)
 */
function getSheetsClient() {
  const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(
    /\\n/g,
    "\n",
  );

  if (!clientEmail || !privateKey) {
    throw new Error(
      "Konfigurasi GOOGLE_SERVICE_ACCOUNT_EMAIL atau GOOGLE_PRIVATE_KEY belum diisi.",
    );
  }

  const auth = new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

/**
 * Mengonversi index kolom (0-based) ke huruf kolom Excel/Sheets (A, B, ..., Z, AA)
 */
function colIndexToLetter(index) {
  let letter = "";
  let temp = index + 1;
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

/**
 * Membersihkan nilai angka dari sel (menangani string format Rupiah / notasi saintifik)
 */
function parseNumber(val) {
  if (typeof val === "number") return val;
  if (!val) return 0;
  const cleaned = String(val).replace(/[^0-9.-]/g, "");
  return Number(cleaned) || 0;
}

/**
 * Anchor-Based Parser untuk sheet D-[BULAN]
 * Mencari header tabel secara dinamis agar tahan terhadap pergeseran baris antar-bulan.
 */
function parseDashboardSheet(rows, sheetName) {
  const incomes = [];
  const budgets = [];
  const DEBT_KEYWORDS = [
    "paylater",
    "hutang",
    "cicilan",
    "service",
    "perjalanan pulang",
    "pinjaman",
  ];
  const SAVING_KEYWORDS = [
    "rdpu",
    "reksadana",
    "tabungan",
    "nabung",
    "emas",
    "saham",
    "dana darurat",
  ];

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || "").trim();

      // 1. Deteksi blok Sumber Pendapatan (Header: Sumber | Tipe | Jumlah)
      if (cell === "Sumber" && String(row[c + 1] || "").trim() === "Tipe") {
        let ptr = r + 1;
        while (ptr < rows.length) {
          const itemRow = rows[ptr] || [];
          const source = String(itemRow[c] || "").trim();
          const type = String(itemRow[c + 1] || "").trim();
          const amount = parseNumber(itemRow[c + 2]);

          if (
            !source ||
            source === "TOTAL" ||
            source === "Keterangan" ||
            source.startsWith("Total")
          )
            break;
          if (amount > 0 || type) {
            incomes.push({
              id: `inc-${sheetName}-${ptr + 1}`,
              source,
              type: type || "Utama",
              amount,
              sheetRow: ptr + 1,
              sheetColStart: colIndexToLetter(c),
            });
          }
          ptr++;
        }
      }

      // 2. Deteksi blok Anggaran / Saving (Header: Keterangan | Prioritas | Rencana | Aktual)
      if (
        cell === "Keterangan" &&
        String(row[c + 1] || "").trim() === "Prioritas" &&
        String(row[c + 2] || "").trim() === "Rencana"
      ) {
        let ptr = r + 1;
        while (ptr < rows.length) {
          const itemRow = rows[ptr] || [];
          const name = String(itemRow[c] || "").trim();
          const priority = String(itemRow[c + 1] || "").trim();
          const planned = parseNumber(itemRow[c + 2]);
          const actual = parseNumber(itemRow[c + 3]);

          if (
            !name ||
            name === "Keterangan" ||
            name === "TOTAL" ||
            name === "Pendapatan"
          )
            break;
          // Lewati baris agregat otomatis yang ditarik dari T-[BULAN]
          if (
            name.toLowerCase().startsWith("total biaya jajan") ||
            name.toLowerCase().startsWith("total pengeluaran sehari")
          ) {
            ptr++;
            continue;
          }

          const lowerName = name.toLowerCase();
          let category = "FIXED";
          if (SAVING_KEYWORDS.some((kw) => lowerName.includes(kw))) {
            category = "SAVING";
          } else if (DEBT_KEYWORDS.some((kw) => lowerName.includes(kw))) {
            category = "DEBT";
          }

          budgets.push({
            id: `bdg-${sheetName}-${ptr + 1}`,
            category,
            name,
            priority: priority || "High",
            planned,
            actual,
            sheetRow: ptr + 1,
            actualCell: `${colIndexToLetter(c + 3)}${ptr + 1}`,
          });
          ptr++;
        }
      }
    }
  }

  return { incomes, budgets };
}

/**
 * Dual Sub-Ledger Parser untuk sheet T-[BULAN]
 * Mendeteksi tabel "Sehari-Hari" (Tanggal | Keterangan | Jumlah | Harga | ... | Total)
 * dan tabel "Jajanan / Nongkrong" (Tanggal | Jajanan / Nongkrong | Uang yang Dihabiskan | ... | Keterangan)
 */
function parseTransactionSheet(rows, sheetName) {
  const transactions = [];
  const tableAnchors = {
    DAILY: null,
    SNACK: null,
  };

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || "").trim();

      // Anchor Tabel 1: Pengeluaran Sehari-Hari
      if (
        cell === "Tanggal" &&
        String(row[c + 1] || "").trim() === "Keterangan" &&
        String(row[c + 2] || "").trim() === "Jumlah"
      ) {
        tableAnchors.DAILY = { headerRow: r + 1, colStart: c };
        let ptr = r + 1;
        while (ptr < rows.length) {
          const tRow = rows[ptr] || [];
          const dateVal = String(tRow[c] || "").trim();
          const descVal = String(tRow[c + 1] || "").trim();
          // Berhenti jika bertemu header tabel lain
          if (dateVal === "Tanggal") break;

          if (dateVal && descVal) {
            const qty = parseNumber(tRow[c + 2]) || 1;
            const unitPrice = parseNumber(tRow[c + 3]);
            const total = parseNumber(tRow[c + 5]) || qty * unitPrice;
            const noteMatch = descVal.match(/\[(Suami|Istri)\]/i);

            transactions.push({
              id: `trx-${sheetName}-D-${ptr + 1}`,
              subLedger: "DAILY",
              date: dateVal,
              desc: descVal.replace(/\s*\[(Suami|Istri)\]/gi, "").trim(),
              qty,
              unitPrice,
              total,
              note: noteMatch ? `[${noteMatch[1]}]` : "",
              by: noteMatch
                ? noteMatch[1].toLowerCase() === "istri"
                  ? "Istri"
                  : "Suami"
                : "Suami",
              sheetRow: ptr + 1,
            });
          } else if (
            !dateVal &&
            !descVal &&
            !tableAnchors.DAILY.firstEmptyRow
          ) {
            tableAnchors.DAILY.firstEmptyRow = ptr + 1;
          }
          ptr++;
        }
      }

      // Anchor Tabel 2: Jajanan / Nongkrong
      if (
        cell === "Tanggal" &&
        String(row[c + 1] || "")
          .trim()
          .includes("Jajanan")
      ) {
        tableAnchors.SNACK = { headerRow: r + 1, colStart: c };
        let ptr = r + 1;
        while (ptr < rows.length) {
          const tRow = rows[ptr] || [];
          const dateVal = String(tRow[c] || "").trim();
          const descVal = String(tRow[c + 1] || "").trim();
          if (dateVal === "Tanggal") break;

          if (dateVal && descVal) {
            const total = parseNumber(tRow[c + 2]);
            const noteVal = String(tRow[c + 4] || "").trim();
            const combinedText = `${descVal} ${noteVal}`;
            const noteMatch = combinedText.match(/\[(Suami|Istri)\]/i);

            transactions.push({
              id: `trx-${sheetName}-S-${ptr + 1}`,
              subLedger: "SNACK",
              date: dateVal,
              desc: descVal,
              qty: 1,
              unitPrice: total,
              total,
              note: noteVal || (noteMatch ? `[${noteMatch[1]}]` : ""),
              by: noteMatch
                ? noteMatch[1].toLowerCase() === "istri"
                  ? "Istri"
                  : "Suami"
                : "Suami",
              sheetRow: ptr + 1,
            });
          } else if (
            !dateVal &&
            !descVal &&
            !tableAnchors.SNACK.firstEmptyRow
          ) {
            tableAnchors.SNACK.firstEmptyRow = ptr + 1;
          }
          ptr++;
        }
      }
    }
  }

  return { transactions, tableAnchors };
}

/**
 * Memastikan sheet D-[BULAN] dan T-[BULAN] tersedia; jika belum ada, buat template standar baru.
 */
async function ensurePeriodSheetsExist(sheets, spreadsheetId, monthCode) {
  const dName = `D-${monthCode}`;
  const tName = `T-${monthCode}`;

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const existingTitles = new Set(
    meta.data.sheets.map((s) => s.properties.title),
  );
  const requests = [];

  if (!existingTitles.has(dName)) {
    requests.push({ addSheet: { properties: { title: dName } } });
  }
  if (!existingTitles.has(tName)) {
    requests.push({ addSheet: { properties: { title: tName } } });
  }

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });

    // Isi struktur header standar pada sheet baru
    if (!existingTitles.has(tName)) {
      const emptyDailyRows = Array.from({ length: 60 }, (_, idx) => {
        const rNum = idx + 2;
        return [
          "",
          "",
          "",
          "",
          "",
          `=IF(AND(C${rNum}<>"",D${rNum}<>""),C${rNum}*D${rNum},0)`,
        ];
      });
      const emptySnackRows = Array.from({ length: 60 }, () => [
        "",
        "",
        0,
        "",
        "",
      ]);

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            {
              range: `${tName}!A1:F61`,
              values: [
                ["Tanggal", "Keterangan", "Jumlah", "Harga", "", "Total"],
                ...emptyDailyRows,
              ],
            },
            {
              range: `${tName}!A64:E124`,
              values: [
                [
                  "Tanggal",
                  "Jajanan / Nongkrong",
                  "Uang yang Dihabiskan",
                  "",
                  "Keterangan",
                ],
                ...emptySnackRows,
              ],
            },
          ],
        },
      });
    }

    if (!existingTitles.has(dName)) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${dName}!A1:E25`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            ["Sumber", "Tipe", "Jumlah", "", ""],
            ["Gaji", "Utama", 0, "", ""],
            ["", "", "", "", ""],
            ["Keterangan", "Prioritas", "Rencana", "Aktual", ""],
            [
              "Total Biaya Jajan & Nongkrong",
              "High",
              "",
              `=SUM('${tName}'!C65:C124)`,
              "",
            ],
            [
              "Total Pengeluaran Sehari-Hari Bulan ini",
              "High",
              "",
              `=SUM('${tName}'!F2:F61)`,
              "",
            ],
            ["", "", "", "", ""],
            ["Keterangan", "Prioritas", "Rencana", "Aktual", ""],
            ["Kost", "High", 350000, 0, ""],
            ["Makan", "High", 450000, 0, ""],
            ["Ortu", "High", 300000, 0, ""],
            ["", "", "", "", ""],
            ["Keterangan", "Prioritas", "Rencana", "Aktual", ""],
            ["Paylater", "High", 0, 0, ""],
            ["", "", "", "", ""],
            ["Keterangan", "Prioritas", "Rencana", "Aktual", ""],
            ["RDPU", "High", 2500000, 0, ""],
          ],
        },
      });
    }
  }

  return { dName, tName };
}

/**
 * ============================================================================
 * MAIN SERVERLESS HANDLER (GET = Pull Sync, POST = Push Sync)
 * ============================================================================
 */
export default async function handler(req, res) {
  // CORS Headers agar fleksibel diakses dari PWA / Frontend
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "Content-Type, x-household-key",
  );

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // Keamanan Zero-Login: Validasi Secret Key Keluarga
  const expectedKey = process.env.HOUSEHOLD_SECRET_KEY;
  const providedKey = req.headers["x-household-key"] || req.query.key;
  if (expectedKey && providedKey !== expectedKey) {
    return res
      .status(401)
      .json({ error: "Unauthorized: Kunci akses keluarga tidak valid." });
  }

  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  if (!spreadsheetId) {
    return res
      .status(500)
      .json({ error: "GOOGLE_SPREADSHEET_ID belum dikonfigurasi." });
  }

  try {
    const sheets = getSheetsClient();
    const period = String(req.query.period || req.body?.period || "2026-04");
    const monthNum = period.split("-")[1] || "04";
    const monthCode = MONTH_CODES[monthNum] || "APR";

    // ------------------------------------------------------------------------
    // GET: PULL DATA DARI GOOGLE SHEETS -> WEB
    // ------------------------------------------------------------------------
    if (req.method === "GET") {
      const { dName, tName } = await ensurePeriodSheetsExist(
        sheets,
        spreadsheetId,
        monthCode,
      );

      const batchRes = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges: [`${dName}!A1:J50`, `${tName}!A1:H150`],
        valueRenderOption: "UNFORMATTED_VALUE",
      });

      const dRows = batchRes.data.valueRanges?.[0]?.values || [];
      const tRows = batchRes.data.valueRanges?.[1]?.values || [];

      const { incomes, budgets } = parseDashboardSheet(dRows, dName);
      const { transactions } = parseTransactionSheet(tRows, tName);

      return res.status(200).json({
        status: "synced",
        period,
        sheets: { dashboard: dName, transaction: tName },
        syncedAt: new Date().toISOString(),
        data: { incomes, budgets, transactions },
      });
    }

    // ------------------------------------------------------------------------
    // POST: PUSH PERUBAHAN DARI WEB -> GOOGLE SHEETS
    // ------------------------------------------------------------------------
    if (req.method === "POST") {
      const { action, payload } = req.body || {};
      const targetPeriod = payload?.date ? payload.date.slice(0, 7) : period;
      const targetMonthCode =
        MONTH_CODES[targetPeriod.split("-")[1]] || monthCode;
      const { dName, tName } = await ensurePeriodSheetsExist(
        sheets,
        spreadsheetId,
        targetMonthCode,
      );

      // ACTION 1: Tambah Transaksi Baru ke T-[BULAN] (Safe Row Append di atas baris 0)
      if (action === "ADD_TRANSACTION") {
        const tGet = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tName}!A1:H150`,
          valueRenderOption: "UNFORMATTED_VALUE",
        });
        const tRows = tGet.data.values || [];
        const { tableAnchors } = parseTransactionSheet(tRows, tName);

        const isDaily = payload.subLedger === "DAILY";
        const anchor = isDaily
          ? tableAnchors.DAILY || {
              headerRow: 1,
              colStart: 0,
              firstEmptyRow: 2,
            }
          : tableAnchors.SNACK || {
              headerRow: 64,
              colStart: 0,
              firstEmptyRow: 65,
            };

        const targetRow = anchor.firstEmptyRow || anchor.headerRow + 1;
        const startCol = colIndexToLetter(anchor.colStart);

        if (isDaily) {
          const endCol = colIndexToLetter(anchor.colStart + 5);
          const descWithTag =
            payload.note && !payload.desc.includes("[")
              ? `${payload.desc} ${payload.note}`
              : payload.desc;

          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${tName}!${startCol}${targetRow}:${endCol}${targetRow}`,
            valueInputOption: "USER_ENTERED",
            requestBody: {
              values: [
                [
                  payload.date,
                  descWithTag,
                  Number(payload.qty) || 1,
                  Number(payload.unitPrice) || 0,
                  "",
                  `=C${targetRow}*D${targetRow}`,
                ],
              ],
            },
          });
        } else {
          const endCol = colIndexToLetter(anchor.colStart + 4);
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${tName}!${startCol}${targetRow}:${endCol}${targetRow}`,
            valueInputOption: "USER_ENTERED",
            requestBody: {
              values: [
                [
                  payload.date,
                  payload.desc,
                  Number(payload.total) || 0,
                  "",
                  payload.note || `[${payload.by || "Suami"}]`,
                ],
              ],
            },
          });
        }

        return res.status(200).json({
          status: "pushed",
          sheet: tName,
          row: targetRow,
          message: `Transaksi berhasil ditulis ke ${tName} baris ${targetRow}`,
        });
      }

      // ACTION 2: Realisasi / Update Nilai Aktual Pos Anggaran di D-[BULAN]
      if (action === "UPDATE_BUDGET_ACTUAL") {
        if (!payload?.actualCell) {
          return res
            .status(400)
            .json({ error: "Koordinat actualCell wajib disertakan." });
        }
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${dName}!${payload.actualCell}`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: [[Number(payload.actual) || 0]],
          },
        });

        return res.status(200).json({
          status: "pushed",
          sheet: dName,
          cell: payload.actualCell,
          message: `Sel ${dName}!${payload.actualCell} diperbarui.`,
        });
      }

      // ACTION 3: Provisioning Bulan Baru (D-[BULAN] & T-[BULAN])
      if (action === "PROVISION_PERIOD") {
        return res.status(200).json({
          status: "provisioned",
          sheets: { dashboard: dName, transaction: tName },
          message: `Tab ${dName} & ${tName} siap digunakan.`,
        });
      }

      return res
        .status(400)
        .json({ error: `Action '${action}' tidak dikenali.` });
    }

    return res.status(405).json({ error: "Method Not Allowed" });
  } catch (err) {
    console.error("[Sheets Sync Error]:", err);
    return res.status(500).json({
      error: "Gagal menyinkronkan dengan Google Sheets",
      details: err.message,
    });
  }
}
