import { google } from "googleapis";

const SPREADSHEET_ID = "1_isC8_KQD0Xs9CiWK9aFHDN-Zp3xXDQQH_bhisBviFA";

const HEADERS = {
  Productos:    ["id", "nombre"],
  Temperaturas: ["productoId", "temp", "tiempo", "indicaciones"],
  Precios:      ["productoId", "precioNormal", "precioMayoreo", "cantMayoreo", "precioSuperMayoreo", "cantSuperMayoreo", "observacion"],
  Margen:       ["productoId", "costoProducto", "dtf", "tinta", "papel", "otros"],
  Ingresos:     ["id", "fecha", "codigoProducto", "nombreProducto", "cliente", "cantidad", "precioUnitario", "precioFinal", "estatus", "pago", "observacion"],
  Gastos:       ["id", "fecha", "codigoProducto", "nombreProducto", "proveedor", "cantidad", "precioUnitario", "precioFinal", "estatus", "pago", "observacion"],
  Clientes:     ["id", "nombre", "telefono", "email", "observacion"],
  Proveedores:  ["id", "nombre", "telefono", "email", "observacion"],
};

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
}

const NUMERIC_FIELDS = new Set(["temp","tiempo","precioNormal","precioMayoreo","cantMayoreo","precioSuperMayoreo","cantSuperMayoreo","costoProducto","dtf","tinta","papel","otros","cantidad","precioUnitario","precioFinal"]);

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const val = row[i] ?? "";
    obj[h] = NUMERIC_FIELDS.has(h) ? (val === "" ? 0 : Number(val) || 0) : val;
  });
  return obj;
}

function objToRow(headers, obj) {
  return headers.map(h => obj[h] ?? "");
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: "v4", auth });
    const { sheet, action } = req.query;

    if (!sheet || !HEADERS[sheet]) {
      return res.status(400).json({ error: "Hoja no válida" });
    }

    const headers = HEADERS[sheet];

    if (req.method === "GET") {
      if (action === "read") {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheet}!A1:Z`,
        });
        const rows = response.data.values || [];
        if (rows.length <= 1) return res.json([]);
        const data = rows.slice(1).map(row => rowToObj(headers, row));
        return res.json(data);
      }
    }

    if (req.method === "POST") {
      const { data } = req.body;

      if (action === "write") {
        const rows = [headers, ...data.map(obj => objToRow(headers, obj))];
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheet}!A1`,
          valueInputOption: "RAW",
          requestBody: { values: rows },
        });
        return res.json({ ok: true });
      }

      if (action === "append") {
        const rows = data.map(obj => objToRow(headers, obj));
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheet}!A1`,
          valueInputOption: "RAW",
          insertDataOption: "INSERT_ROWS",
          requestBody: { values: rows },
        });
        return res.json({ ok: true });
      }
    }

    return res.status(400).json({ error: "Acción no válida" });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
