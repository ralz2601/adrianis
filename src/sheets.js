// Hook para sincronizar datos con Google Sheets
const API = "/api/sheets";

export async function leerSheet(sheet) {
  try {
    const res = await fetch(`${API}?sheet=${sheet}&action=read`);
    if (!res.ok) throw new Error("Error al leer");
    return await res.json();
  } catch (err) {
    console.error(`Error leyendo ${sheet}:`, err);
    return null;
  }
}

export async function escribirSheet(sheet, data) {
  try {
    const res = await fetch(`${API}?sheet=${sheet}&action=write`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ data }),
    });
    if (!res.ok) throw new Error("Error al escribir");
    return true;
  } catch (err) {
    console.error(`Error escribiendo ${sheet}:`, err);
    return false;
  }
}

export async function agregarSheet(sheet, data) {
  try {
    const res = await fetch(`${API}?sheet=${sheet}&action=append`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ data }),
    });
    if (!res.ok) throw new Error("Error al agregar");
    return true;
  } catch (err) {
    console.error(`Error agregando a ${sheet}:`, err);
    return false;
  }
}
