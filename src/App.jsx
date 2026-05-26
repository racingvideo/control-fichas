import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";

const FORMATIONS = {
  "1-4-3-3": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Lateral derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Lateral izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 3 }, { pos: "Defensa Central", score: 5 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Enganche", score: 3 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Enganche", score: 3 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Ataque", label: "Extremo derecho", accepts: [{ pos: "Extremo Derecho", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Izquierdo", score: 4 }, { pos: "Delantero Centro", score: 6 }] },
    { line: "Ataque", label: "Delantero centro", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] },
    { line: "Ataque", label: "Extremo izquierdo", accepts: [{ pos: "Extremo Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Derecho", score: 4 }, { pos: "Delantero Centro", score: 6 }] }
  ],
  "1-4-2-3-1": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Lateral derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Lateral izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Volantes defensivos", label: "Doble 5", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 3 }, { pos: "Defensa Central", score: 5 }] },
    { line: "Volantes defensivos", label: "Doble 5", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 3 }, { pos: "Defensa Central", score: 5 }] },
    { line: "Tres cuartos", label: "Extremo derecho", accepts: [{ pos: "Extremo Derecho", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Izquierdo", score: 4 }] },
    { line: "Tres cuartos", label: "Enganche", accepts: [{ pos: "Enganche", score: 0 }, { pos: "Volante Interno", score: 2 }, { pos: "Volante Central (5)", score: 3 }, { pos: "Volante Extremo", score: 3 }, { pos: "Delantero Centro", score: 5 }] },
    { line: "Tres cuartos", label: "Extremo izquierdo", accepts: [{ pos: "Extremo Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Derecho", score: 4 }] },
    { line: "Ataque", label: "Delantero centro", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] }
  ],
  "1-4-4-2": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Lateral derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Lateral izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Volante Extremo", score: 5 }] },
    { line: "Mediocampo", label: "Volante derecho", accepts: [{ pos: "Volante Extremo", score: 0 }, { pos: "Extremo Derecho", score: 1 }, { pos: "Lateral Derecho", score: 3 }, { pos: "Volante Interno", score: 4 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 2 }, { pos: "Enganche", score: 4 }, { pos: "Defensa Central", score: 5 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Enganche", score: 3 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Mediocampo", label: "Volante izquierdo", accepts: [{ pos: "Volante Extremo", score: 0 }, { pos: "Extremo Izquierdo", score: 1 }, { pos: "Lateral Izquierdo", score: 3 }, { pos: "Volante Interno", score: 4 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] }
  ],
  "1-3-5-2": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }] },
    { line: "Mediocampo", label: "Carrilero derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Derecho", score: 3 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 2 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Enganche", score: 1 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Enganche", score: 1 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Mediocampo", label: "Carrilero izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Izquierdo", score: 3 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] }
  ],
  "1-5-3-2": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Lateral derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Volante Extremo", score: 4 }, { pos: "Extremo Derecho", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Izquierdo", score: 4 }] },
    { line: "Defensa", label: "Lateral izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 4 }, { pos: "Extremo Izquierdo", score: 5 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 2 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Enganche", score: 1 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Mediocampo", label: "Interno", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Enganche", score: 1 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Volante Extremo", score: 4 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] },
    { line: "Ataque", label: "Delantero", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] }
  ],
  "1-3-4-3": [
    { line: "Arquero", label: "Arquero", accepts: [{ pos: "Arquero", score: 0 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Volante Central (5)", score: 5 }] },
    { line: "Defensa", label: "Zaguero", accepts: [{ pos: "Defensa Central", score: 0 }, { pos: "Lateral Derecho", score: 4 }, { pos: "Lateral Izquierdo", score: 4 }] },
    { line: "Mediocampo", label: "Carrilero derecho", accepts: [{ pos: "Lateral Derecho", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Derecho", score: 3 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Central (5)", score: 0 }, { pos: "Volante Interno", score: 2 }] },
    { line: "Mediocampo", label: "Volante central", accepts: [{ pos: "Volante Interno", score: 0 }, { pos: "Volante Central (5)", score: 2 }, { pos: "Enganche", score: 3 }] },
    { line: "Mediocampo", label: "Carrilero izquierdo", accepts: [{ pos: "Lateral Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Izquierdo", score: 3 }] },
    { line: "Ataque", label: "Extremo derecho", accepts: [{ pos: "Extremo Derecho", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Izquierdo", score: 4 }] },
    { line: "Ataque", label: "Delantero centro", accepts: [{ pos: "Delantero Centro", score: 0 }, { pos: "Extremo Derecho", score: 5 }, { pos: "Extremo Izquierdo", score: 5 }, { pos: "Enganche", score: 6 }] },
    { line: "Ataque", label: "Extremo izquierdo", accepts: [{ pos: "Extremo Izquierdo", score: 0 }, { pos: "Volante Extremo", score: 1 }, { pos: "Extremo Derecho", score: 4 }] }
  ]
};

const LINE_ORDER = ["Arquero", "Defensa", "Volantes defensivos", "Mediocampo", "Tres cuartos", "Ataque"];
const EMPTY_DUPLICATE_SUMMARY = { samePlayer: [], sameNameDifferentDate: [] };

function cleanText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizeText(value) {
  return cleanText(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function getColumnValue(row, aliases) {
  const normalizedRow = {};

  Object.keys(row).forEach((key) => {
    normalizedRow[normalizeText(key)] = row[key];
  });

  for (const alias of aliases) {
    const normalizedAlias = normalizeText(alias);
    if (Object.prototype.hasOwnProperty.call(normalizedRow, normalizedAlias)) {
      return normalizedRow[normalizedAlias];
    }
  }

  return "";
}

function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number") {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    const date = new Date(utcValue * 1000);
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  }

  const text = cleanText(value);
  const match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) return date;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function padNumber(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  if (!date) return "";
  return `${padNumber(date.getDate())}/${padNumber(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function formatDateKey(date) {
  if (!date) return "";
  return `${date.getFullYear()}-${padNumber(date.getMonth() + 1)}-${padNumber(date.getDate())}`;
}

function canonicalizePosition(value) {
  const text = normalizeText(value);

  if (!text || text === "none" || text === "posicion en cancha") return "";

  if (text.includes("arquero") || text.includes("golero") || text.includes("portero")) return "Arquero";
  if (text.includes("defensa central") || text.includes("zaguero") || text.includes("central")) return "Defensa Central";

  if (text.includes("lateral derecho") || text === "ld" || text.includes("marcador derecho")) return "Lateral Derecho";
  if (text.includes("lateral izquierdo") || text === "li" || text.includes("marcador izquierdo")) return "Lateral Izquierdo";

  if (text.includes("volante central") || text.includes("mediocentro") || text.includes("medio centro") || text === "5" || text.includes("(5)")) return "Volante Central (5)";
  if (text.includes("volante interno") || text.includes("interno")) return "Volante Interno";
  if (text.includes("volante externo") || text.includes("volante extremo") || text.includes("carrilero")) return "Volante Extremo";

  if (text.includes("extremo derecho") || text === "ed" || text.includes("punta derecha")) return "Extremo Derecho";
  if (text.includes("extremo izquierdo") || text === "ei" || text.includes("punta izquierda")) return "Extremo Izquierdo";

  if (text.includes("delantero centro") || text.includes("delantero central") || text.includes("centro delantero") || text.includes("centrodelantero") || text === "9") return "Delantero Centro";
  if (text.includes("enganche") || text.includes("media punta") || text.includes("mediapunta") || text.includes("10")) return "Enganche";

  return cleanText(value);
}

function splitPositions(value) {
  const positions = cleanText(value)
    .split(",")
    .map((item) => canonicalizePosition(item))
    .filter(Boolean);

  return [...new Set(positions)];
}

function getCategory(date) {
  if (!date) return "Fecha inválida";
  const year = date.getFullYear();
  if (year < 2000 || year > 2020) return "Fecha inválida";
  return String(year);
}

function getIdentityKey(player) {
  const nameKey = normalizeText(player.name);
  const dateKey = formatDateKey(player.birthDate);
  if (!nameKey || !dateKey) return `fila-${player.rowNumber}`;
  return `${nameKey}|${dateKey}`;
}

function normalizeRow(row, index) {
  const name = cleanText(getColumnValue(row, [
    "Nombre completo",
    "Nombre Completo",
    "NOMBRE COMPLETO",
    "Nombre y apellido",
    "Nombre y Apellido",
    "NOMBRE Y APELLIDO",
    "Jugador",
    "JUGADOR",
    "Nombre"
  ]));

  const birthDate = parseExcelDate(getColumnValue(row, [
    "Fecha de nacimiento",
    "Fecha de Nacimiento",
    "FECHA DE NACIMIENTO",
    "Fecha nacimiento",
    "Fecha Nacimiento",
    "FECHA NACIMIENTO",
    "Nacimiento",
    "NACIMIENTO",
    "Fecha nac.",
    "Fecha nac",
    "Fecha de nac."
  ]));

  const positions = splitPositions(getColumnValue(row, [
    "Posición en cancha",
    "Posicion en cancha",
    "Posición en Cancha",
    "Posicion en Cancha",
    "POSICIÓN EN CANCHA",
    "POSICION EN CANCHA",
    "Posición",
    "Posicion",
    "POSICIÓN",
    "POSICION",
    "Puesto",
    "PUESTO"
  ]));

  const leg = cleanText(getColumnValue(row, [
    "Pierna Hábil",
    "Pierna hábil",
    "Pierna habil",
    "Pierna Habil",
    "PIERNA HÁBIL",
    "PIERNA HABIL",
    "Pierna",
    "PIERNA",
    "Perfil",
    "PERFIL"
  ]));

  const club = cleanText(getColumnValue(row, [
    "Club actual o último club",
    "Club actual o último Club",
    "Club actual o ultimo club",
    "Club actual o ultimo Club",
    "CLUB ACTUAL O ÚLTIMO CLUB",
    "CLUB ACTUAL O ULTIMO CLUB",
    "Club actual",
    "CLUB ACTUAL",
    "Último club",
    "Ultimo club",
    "ULTIMO CLUB"
  ]));

  const background = cleanText(getColumnValue(row, [
    "Antecedentes en Auf o Selección OFI",
    "Antecedentes en AUF o Selección OFI",
    "Antecedentes en Auf o Seleccion OFI",
    "Antecedentes en AUF o Seleccion OFI",
    "ANTECEDENTES EN AUF O SELECCIÓN OFI",
    "ANTECEDENTES EN AUF O SELECCION OFI",
    "Antecedentes",
    "ANTECEDENTES",
    "Antecedentes AUF",
    "Selección OFI",
    "Seleccion OFI"
  ]));

  const passStatus = cleanText(getColumnValue(row, [
    "Calidad del pase",
    "Calidad del Pase",
    "CALIDAD DEL PASE",
    "Pase",
    "PASE",
    "Condición del pase",
    "Condicion del pase",
    "CONDICION DEL PASE"
  ]));

  const category = getCategory(birthDate);
  const rowNumber = index + 2;

  return {
    id: `fila-${rowNumber}`,
    rowNumber,
    identityKey: "",
    name,
    nameKey: normalizeText(name),
    positions,
    positionText: positions.join(", "),
    leg,
    birthDate,
    birthDateText: formatDate(birthDate),
    birthDateKey: formatDateKey(birthDate),
    category,
    club,
    background,
    passStatus,
    duplicateRows: [],
    isManual: false,
    hasValidData: Boolean(name) && positions.length > 0 && category !== "Fecha inválida"
  };
}

function makeManualPlayer(name, teamIndex, pickIndex, previousPlayer = null) {
  const cleanName = cleanText(name);

  return {
    id: previousPlayer?.isManual ? previousPlayer.id : `manual-${teamIndex}-${pickIndex}-${Date.now()}`,
    rowNumber: "",
    identityKey: "",
    name: cleanName,
    nameKey: normalizeText(cleanName),
    positions: [],
    positionText: "Agregado manualmente",
    leg: "",
    birthDate: null,
    birthDateText: "",
    birthDateKey: "",
    category: "Manual",
    club: "",
    background: "",
    passStatus: "",
    duplicateRows: [],
    isManual: true,
    hasValidData: Boolean(cleanName)
  };
}

function mergePlayerData(basePlayer, repeatedPlayer) {
  const positions = [...new Set([...basePlayer.positions, ...repeatedPlayer.positions])];
  basePlayer.positions = positions;
  basePlayer.positionText = positions.join(", ");
  basePlayer.duplicateRows.push(repeatedPlayer);
  basePlayer.hasValidData = Boolean(basePlayer.name) && positions.length > 0 && basePlayer.category !== "Fecha inválida";
}

function processRows(rows) {
  const rawPlayers = rows
    .map(normalizeRow)
    .filter((player) => player.name && normalizeText(player.name) !== "nombre completo");

  const byIdentity = new Map();
  const uniquePlayers = [];

  rawPlayers.forEach((player) => {
    const identityKey = getIdentityKey(player);
    player.identityKey = identityKey;

    if (byIdentity.has(identityKey)) {
      mergePlayerData(byIdentity.get(identityKey), player);
    } else {
      const playerWithIdentity = { ...player, id: identityKey };
      byIdentity.set(identityKey, playerWithIdentity);
      uniquePlayers.push(playerWithIdentity);
    }
  });

  const samePlayer = uniquePlayers
    .filter((player) => player.duplicateRows.length > 0)
    .map((player) => ({
      name: player.name,
      positionText: player.positionText || "Sin posición",
      birthDateText: player.birthDateText || "Sin fecha",
      count: player.duplicateRows.length + 1,
      rows: [player.rowNumber, ...player.duplicateRows.map((item) => item.rowNumber)]
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const byName = new Map();
  uniquePlayers.forEach((player) => {
    if (!player.nameKey) return;
    if (!byName.has(player.nameKey)) byName.set(player.nameKey, []);
    byName.get(player.nameKey).push(player);
  });

  const sameNameDifferentDate = [...byName.values()]
    .filter((group) => group.length > 1)
    .map((group) => ({
      name: group[0].name,
      count: group.length,
      entries: group
        .map((player) => ({
          positionText: player.positionText || "Sin posición",
          birthDateText: player.birthDateText || "Sin fecha",
          rowNumber: player.rowNumber
        }))
        .sort((a, b) => a.birthDateText.localeCompare(b.birthDateText))
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  return {
    rawTotal: rawPlayers.length,
    players: uniquePlayers,
    duplicateSummary: { samePlayer, sameNameDifferentDate }
  };
}

function playerRoleScore(player, role) {
  if (!player || player.isManual) return null;

  const scores = role.accepts
    .filter((rule) => player.positions.includes(rule.pos))
    .map((rule) => rule.score);

  if (!scores.length) return null;
  return Math.min(...scores);
}

function isManualAdjustment(pick) {
  if (!pick.player || pick.player.isManual) return false;
  return playerRoleScore(pick.player, pick.role) === null;
}

function getPlayerDescription(pick) {
  if (!pick.player) return "";
  if (pick.player.isManual) return "Jugador agregado manualmente";

  const description = `${pick.player.positionText} · ${pick.player.leg || "Sin pierna"}`;
  return isManualAdjustment(pick) ? `${description} · Ajuste manual` : description;
}

function getLegBonus(player, role) {
  const leg = normalizeText(player.leg);
  const label = normalizeText(role.label);
  if (!leg) return 0;
  if (label.includes("izquierdo") && leg.includes("izquierda")) return -0.4;
  if (label.includes("derecho") && leg.includes("derecha")) return -0.4;
  if (label.includes("izquierdo") && leg.includes("derecha")) return 0.6;
  if (label.includes("derecho") && leg.includes("izquierda")) return 0.6;
  return 0;
}

function getPlayerFlexibility(player, roles) {
  return roles.reduce((total, role) => total + (playerRoleScore(player, role) === null ? 0 : 1), 0);
}

function makeEmptyPick(role, roleIndex) {
  return { player: null, role, roleIndex, score: null };
}

function recalculateTeamStatus(team) {
  return {
    ...team,
    isPartial: team.picked.filter((item) => item.player).length < team.picked.length
  };
}

function getRoleOrder(remainingPlayers, roles) {
  return roles
    .map((role, roleIndex) => ({
      role,
      roleIndex,
      candidatesCount: remainingPlayers.filter((player) => playerRoleScore(player, role) !== null).length
    }))
    .sort((a, b) => a.candidatesCount - b.candidatesCount || a.roleIndex - b.roleIndex);
}

function getRoleCandidates(remainingPlayers, used, role, roleIndex, roles) {
  return remainingPlayers
    .filter((player) => !used.has(player.id))
    .map((player) => {
      const baseScore = playerRoleScore(player, role);
      if (baseScore === null) return null;
      return {
        player,
        role,
        roleIndex,
        score: baseScore + getLegBonus(player, role),
        flexibility: getPlayerFlexibility(player, roles)
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.score - b.score || a.flexibility - b.flexibility || a.player.rowNumber - b.player.rowNumber);
}

function buildCompleteTeam(remainingPlayers, roles) {
  const roleOrder = getRoleOrder(remainingPlayers, roles);
  if (roleOrder.some((item) => item.candidatesCount === 0)) return null;

  const used = new Set();
  const picked = Array.from({ length: roles.length }, (_, roleIndex) => makeEmptyPick(roles[roleIndex], roleIndex));
  let attempts = 0;
  const maxAttempts = 60000;

  function search(orderIndex) {
    attempts += 1;
    if (attempts > maxAttempts) return false;
    if (orderIndex === roleOrder.length) return true;

    const { role, roleIndex } = roleOrder[orderIndex];
    const candidates = getRoleCandidates(remainingPlayers, used, role, roleIndex, roles);

    for (const candidate of candidates) {
      used.add(candidate.player.id);
      picked[roleIndex] = candidate;
      if (search(orderIndex + 1)) return true;
      used.delete(candidate.player.id);
      picked[roleIndex] = makeEmptyPick(role, roleIndex);
    }

    return false;
  }

  return search(0) ? picked : null;
}

function buildPartialTeam(remainingPlayers, roles) {
  const used = new Set();
  const picked = Array.from({ length: roles.length }, (_, roleIndex) => makeEmptyPick(roles[roleIndex], roleIndex));
  const roleOrder = getRoleOrder(remainingPlayers, roles);

  roleOrder.forEach(({ role, roleIndex }) => {
    const [bestCandidate] = getRoleCandidates(remainingPlayers, used, role, roleIndex, roles);
    if (!bestCandidate) return;
    used.add(bestCandidate.player.id);
    picked[roleIndex] = bestCandidate;
  });

  return picked;
}

function generateMaximumTeams(players, formationName) {
  const roles = FORMATIONS[formationName];
  let remainingPlayers = players.filter((player) => player.hasValidData);
  const teams = [];

  while (remainingPlayers.length > 0) {
    const completeTeam = remainingPlayers.length >= roles.length ? buildCompleteTeam(remainingPlayers, roles) : null;
    const picked = completeTeam || buildPartialTeam(remainingPlayers, roles);
    const usedIds = new Set(picked.filter((item) => item.player).map((item) => item.player.id));

    if (usedIds.size === 0) break;

    teams.push({
      name: `Equipo ${teams.length + 1}`,
      formationName,
      picked,
      isPartial: usedIds.size < roles.length
    });

    remainingPlayers = remainingPlayers.filter((player) => !usedIds.has(player.id));
  }

  return teams;
}

function groupTeam(team) {
  return LINE_ORDER.map((line) => ({
    line,
    items: team.picked.filter((item) => item.role.line === line)
  })).filter((group) => group.items.length);
}

function getVisualLineItems(items) {
  return [...items].reverse();
}

function getSwapOptions(teams, currentTeamIndex, currentPick) {
  if (!currentPick.player) return [];

  const currentLabel = normalizeText(currentPick.role.label);

  return teams.flatMap((team, teamIndex) => {
    if (teamIndex === currentTeamIndex) return [];

    return team.picked
      .map((pick, pickIndex) => ({ pick, pickIndex }))
      .filter(({ pick }) => pick.player && normalizeText(pick.role.label) === currentLabel)
      .map(({ pick, pickIndex }) => ({
        value: `${teamIndex}|${pickIndex}`,
        teamIndex,
        pickIndex,
        teamName: team.name,
        player: pick.player,
        role: pick.role
      }));
  });
}

function getInternalSwapOptions(team, currentPickIndex) {
  return team.picked
    .map((pick, pickIndex) => ({ pick, pickIndex }))
    .filter(({ pick, pickIndex }) => pickIndex !== currentPickIndex && pick.player)
    .map(({ pick, pickIndex }) => ({
      value: String(pickIndex),
      pickIndex,
      player: pick.player,
      role: pick.role
    }));
}

function updatePickPlayer(pick, player) {
  return {
    ...pick,
    player,
    score: player ? playerRoleScore(player, pick.role) : null
  };
}

function sanitizeFileName(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "equipos";
}

function App() {
  const [players, setPlayers] = useState([]);
  const [rawTotal, setRawTotal] = useState(0);
  const [duplicateSummary, setDuplicateSummary] = useState(EMPTY_DUPLICATE_SUMMARY);
  const [fileName, setFileName] = useState("");
  const [selectedCategories, setSelectedCategories] = useState([]);
  const [formation, setFormation] = useState("1-4-3-3");
  const [teams, setTeams] = useState([]);
  const [removedPlayers, setRemovedPlayers] = useState([]);
  const [error, setError] = useState("");
  const [isExporting, setIsExporting] = useState(false);

  const categories = useMemo(() => {
    return [...new Set(players.map((player) => player.category))]
      .filter(Boolean)
      .sort((a, b) => {
        if (a === "Fecha inválida") return 1;
        if (b === "Fecha inválida") return -1;
        return Number(a) - Number(b);
      });
  }, [players]);

  const validCategories = useMemo(() => {
    return categories.filter((item) => item !== "Fecha inválida");
  }, [categories]);

  const selectedCategoryLabel = useMemo(() => {
    return selectedCategories.length ? selectedCategories.join(", ") : "Sin categorías";
  }, [selectedCategories]);

  const filteredPlayers = useMemo(() => {
    return players.filter((player) => selectedCategories.includes(player.category) && player.hasValidData);
  }, [players, selectedCategories]);

  const invalidPlayers = useMemo(() => {
    return players.filter((player) => !player.hasValidData);
  }, [players]);

  const positionCount = useMemo(() => {
    const count = {};
    filteredPlayers.forEach((player) => {
      player.positions.forEach((position) => {
        count[position] = (count[position] || 0) + 1;
      });
    });
    return Object.entries(count).sort((a, b) => b[1] - a[1]);
  }, [filteredPlayers]);

  const usedPlayers = useMemo(() => {
    return new Set(teams.flatMap((team) => team.picked.filter((item) => item.player).map((item) => item.player.id)));
  }, [teams]);

  const manualPlayersCount = useMemo(() => {
    return teams.flatMap((team) => team.picked).filter((item) => item.player?.isManual).length;
  }, [teams]);

  const notUsedPlayers = useMemo(() => {
    return filteredPlayers.filter((player) => !usedPlayers.has(player.id));
  }, [filteredPlayers, usedPlayers]);

  const omittedDuplicatesCount = duplicateSummary.samePlayer.reduce((total, item) => total + item.count - 1, 0);
  const completeTeamsCount = teams.filter((team) => !team.isPartial).length;
  const partialTeamsCount = teams.filter((team) => team.isPartial).length;

  function resetTeamsAndManualChanges() {
    setTeams([]);
    setRemovedPlayers([]);
  }

  function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setError("");
    setPlayers([]);
    setRawTotal(0);
    setDuplicateSummary(EMPTY_DUPLICATE_SUMMARY);
    resetTeamsAndManualChanges();

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: false });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: true });
        const processed = processRows(rows);

        setPlayers(processed.players);
        setRawTotal(processed.rawTotal);
        setDuplicateSummary(processed.duplicateSummary);

        const firstValidCategory = [...new Set(processed.players.map((player) => player.category))]
          .filter((item) => item && item !== "Fecha inválida")
          .sort((a, b) => Number(a) - Number(b))[0];

        setSelectedCategories(firstValidCategory ? [firstValidCategory] : []);
      } catch (err) {
        setError("No se pudo leer el Excel. Revisá que sea un archivo .xlsx válido y que mantenga columnas como nombre, fecha de nacimiento y posición en cancha.");
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function toggleCategory(categoryToToggle) {
    resetTeamsAndManualChanges();
    setError("");
    setSelectedCategories((current) => {
      if (current.includes(categoryToToggle)) {
        return current.filter((item) => item !== categoryToToggle);
      }
      return [...current, categoryToToggle].sort((a, b) => Number(a) - Number(b));
    });
  }

  function selectAllCategories() {
    resetTeamsAndManualChanges();
    setError("");
    setSelectedCategories(validCategories);
  }

  function clearCategories() {
    resetTeamsAndManualChanges();
    setError("");
    setSelectedCategories([]);
  }

  function handleGenerate() {
    setRemovedPlayers([]);

    if (!selectedCategories.length) {
      setTeams([]);
      setError("Seleccioná una o más categorías para generar equipos.");
      return;
    }

    if (!filteredPlayers.length) {
      setTeams([]);
      setError("No hay jugadores válidos para las categorías seleccionadas.");
      return;
    }

    const generatedTeams = generateMaximumTeams(filteredPlayers, formation);
    setTeams(generatedTeams);

    if (!generatedTeams.length) {
      setError("No se pudo ubicar ningún jugador con la formación seleccionada. Revisá si las posiciones del Excel coinciden con las posiciones aceptadas.");
      return;
    }

    setError("");
  }

  function handleSwapPlayers(fromTeamIndex, fromPickIndex, targetValue) {
    if (!targetValue) return;

    const [toTeamIndexText, toPickIndexText] = targetValue.split("|");
    const toTeamIndex = Number(toTeamIndexText);
    const toPickIndex = Number(toPickIndexText);

    if (!Number.isInteger(toTeamIndex) || !Number.isInteger(toPickIndex)) return;

    setTeams((currentTeams) => {
      const nextTeams = currentTeams.map((team) => ({
        ...team,
        picked: team.picked.map((pick) => ({ ...pick }))
      }));

      const fromPick = nextTeams[fromTeamIndex]?.picked[fromPickIndex];
      const toPick = nextTeams[toTeamIndex]?.picked[toPickIndex];

      if (!fromPick?.player || !toPick?.player) return currentTeams;
      if (normalizeText(fromPick.role.label) !== normalizeText(toPick.role.label)) return currentTeams;

      const fromPlayer = fromPick.player;
      const toPlayer = toPick.player;

      nextTeams[fromTeamIndex].picked[fromPickIndex] = updatePickPlayer(fromPick, toPlayer);
      nextTeams[toTeamIndex].picked[toPickIndex] = updatePickPlayer(toPick, fromPlayer);

      nextTeams[fromTeamIndex] = recalculateTeamStatus(nextTeams[fromTeamIndex]);
      nextTeams[toTeamIndex] = recalculateTeamStatus(nextTeams[toTeamIndex]);

      return nextTeams;
    });
  }

  function handleInternalSwapPlayers(teamIndex, fromPickIndex, targetValue) {
    if (!targetValue) return;

    const toPickIndex = Number(targetValue);
    if (!Number.isInteger(toPickIndex)) return;
    if (fromPickIndex === toPickIndex) return;

    setTeams((currentTeams) => {
      const nextTeams = currentTeams.map((team) => ({
        ...team,
        picked: team.picked.map((pick) => ({ ...pick }))
      }));

      const fromPick = nextTeams[teamIndex]?.picked[fromPickIndex];
      const toPick = nextTeams[teamIndex]?.picked[toPickIndex];

      if (!fromPick?.player || !toPick?.player) return currentTeams;

      const fromPlayer = fromPick.player;
      const toPlayer = toPick.player;

      nextTeams[teamIndex].picked[fromPickIndex] = updatePickPlayer(fromPick, toPlayer);
      nextTeams[teamIndex].picked[toPickIndex] = updatePickPlayer(toPick, fromPlayer);
      nextTeams[teamIndex] = recalculateTeamStatus(nextTeams[teamIndex]);

      return nextTeams;
    });
  }

  function handleManualNameChange(teamIndex, pickIndex, value) {
    setTeams((currentTeams) => {
      const nextTeams = currentTeams.map((team) => ({
        ...team,
        picked: team.picked.map((pick) => ({ ...pick }))
      }));

      const pick = nextTeams[teamIndex]?.picked[pickIndex];
      if (!pick) return currentTeams;

      const manualName = cleanText(value);
      if (!manualName) {
        nextTeams[teamIndex].picked[pickIndex] = updatePickPlayer(pick, null);
      } else {
        const manualPlayer = makeManualPlayer(manualName, teamIndex, pickIndex, pick.player);
        nextTeams[teamIndex].picked[pickIndex] = updatePickPlayer(pick, manualPlayer);
      }

      nextTeams[teamIndex] = recalculateTeamStatus(nextTeams[teamIndex]);

      return nextTeams;
    });
  }

  function handleRemovePlayer(teamIndex, pickIndex) {
    setTeams((currentTeams) => {
      const nextTeams = currentTeams.map((team) => ({
        ...team,
        picked: team.picked.map((pick) => ({ ...pick }))
      }));

      const team = nextTeams[teamIndex];
      const pick = team?.picked[pickIndex];

      if (!team || !pick?.player) return currentTeams;

      const removedPlayer = pick.player;

      setRemovedPlayers((currentRemoved) => [
        ...currentRemoved,
        {
          id: `${Date.now()}-${teamIndex}-${pickIndex}-${removedPlayer.id}`,
          name: removedPlayer.name,
          positionText: removedPlayer.positionText || "Sin posición",
          birthDateText: removedPlayer.birthDateText || "",
          teamName: team.name,
          roleLabel: pick.role.label,
          origin: removedPlayer.isManual ? "Manual" : "Excel"
        }
      ]);

      nextTeams[teamIndex].picked[pickIndex] = updatePickPlayer(pick, null);
      nextTeams[teamIndex] = recalculateTeamStatus(nextTeams[teamIndex]);

      return nextTeams;
    });
  }

  async function handleExport(type, singleTeamIndex = null) {
    if (!teams.length) return;

    const teamElements = Array.from(document.querySelectorAll("[data-export-team='true']"));
    if (!teamElements.length) return;

    const selectedElements = singleTeamIndex === null
      ? teamElements.map((element, index) => ({ element, index }))
      : teamElements
          .map((element, index) => ({ element, index }))
          .filter((item) => item.index === singleTeamIndex);

    if (!selectedElements.length) return;

    try {
      setIsExporting(true);
      setError("");
      await new Promise((resolve) => requestAnimationFrame(resolve));
      await new Promise((resolve) => setTimeout(resolve, 50));

      const baseName = sanitizeFileName(`equipos-${selectedCategoryLabel}-${formation}`);

      if (type === "png") {
        for (const item of selectedElements) {
          const canvas = await html2canvas(item.element, {
            scale: 2,
            backgroundColor: "#eef4ef",
            useCORS: true
          });

          const image = canvas.toDataURL("image/png");
          const link = document.createElement("a");
          link.href = image;
          link.download = `${baseName}-equipo-${item.index + 1}.png`;
          link.click();
        }
      }

      if (type === "pdf") {
        const pdf = new jsPDF("p", "mm", "a4");
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 8;
        const availableWidth = pageWidth - margin * 2;
        const availableHeight = pageHeight - margin * 2;

        for (let exportIndex = 0; exportIndex < selectedElements.length; exportIndex += 1) {
          const item = selectedElements[exportIndex];
          const canvas = await html2canvas(item.element, {
            scale: 2,
            backgroundColor: "#eef4ef",
            useCORS: true
          });

          const image = canvas.toDataURL("image/png");
          const imageRatio = canvas.width / canvas.height;
          let imageWidth = availableWidth;
          let imageHeight = imageWidth / imageRatio;

          if (imageHeight > availableHeight) {
            imageHeight = availableHeight;
            imageWidth = imageHeight * imageRatio;
          }

          const x = (pageWidth - imageWidth) / 2;
          const y = (pageHeight - imageHeight) / 2;

          if (exportIndex > 0) pdf.addPage();
          pdf.addImage(image, "PNG", x, y, imageWidth, imageHeight);
        }

        const fileName = singleTeamIndex === null ? `${baseName}.pdf` : `${baseName}-equipo-${singleTeamIndex + 1}.pdf`;
        pdf.save(fileName);
      }
    } catch (err) {
      setError("No se pudo exportar el archivo. Probá nuevamente o exportá menos equipos por vez.");
    } finally {
      setIsExporting(false);
    }
  }

  return (
    <main className={`page ${isExporting ? "exporting" : ""}`}>
      <section className="hero">
        <div>
          <p className="eyebrow">RACING DE MONTEVIDEO SAD</p>
          <h1>Armador automático de 11</h1>
          <p className="subtitle">Cargá el Excel de inscripción, combiná una o varias categorías y generá la mayor cantidad posible de equipos sin repetir jugadores.</p>
        </div>
        <div className="badge">Captación</div>
      </section>

      <section className="panel upload-panel">
        <div>
          <h2>Cargar Excel</h2>
          <p>Usá el archivo original del formulario. La categoría se calcula automáticamente desde la fecha de nacimiento.</p>
        </div>
        <label className="file-button">
          Seleccionar Excel
          <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
        </label>
        {fileName && <span className="file-name">{fileName}</span>}
      </section>

      {error && <div className="alert">{error}</div>}

      {players.length > 0 && (
        <>
          <section className="stats-grid">
            <article className="stat-card">
              <span>Respuestas cargadas</span>
              <strong>{rawTotal}</strong>
            </article>
            <article className="stat-card">
              <span>Jugadores únicos</span>
              <strong>{players.length}</strong>
            </article>
            <article className="stat-card">
              <span>Jugadores seleccionados</span>
              <strong>{filteredPlayers.length}</strong>
            </article>
            <article className="stat-card warning-card">
              <span>Duplicados omitidos</span>
              <strong>{omittedDuplicatesCount}</strong>
            </article>
          </section>

          <section className="panel controls multi-category-controls">
            <div className="control-group category-control">
              <div className="control-heading">
                <label>Categorías a mezclar</label>
                <div className="mini-actions">
                  <button type="button" onClick={selectAllCategories}>Todas</button>
                  <button type="button" onClick={clearCategories}>Limpiar</button>
                </div>
              </div>
              <div className="category-grid">
                {validCategories.map((item) => (
                  <label className={`category-check ${selectedCategories.includes(item) ? "active" : ""}`} key={item}>
                    <input type="checkbox" checked={selectedCategories.includes(item)} onChange={() => toggleCategory(item)} />
                    <span>{item}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className="control-group">
              <label>Formación</label>
              <select value={formation} onChange={(event) => { setFormation(event.target.value); resetTeamsAndManualChanges(); setError(""); }}>
                {Object.keys(FORMATIONS).map((item) => (
                  <option key={item} value={item}>{item}</option>
                ))}
              </select>
            </div>

            <button className="primary-button" onClick={handleGenerate}>Generar equipos</button>
          </section>

          <section className="layout-grid">
            <aside className="panel side-panel">
              <h2>Resumen categorías {selectedCategoryLabel}</h2>
              <p className="small-text side-note">La app usa jugadores únicos, mezcla las categorías elegidas y no repite jugadores entre equipos.</p>
              <div className="position-list">
                {positionCount.map(([position, count]) => (
                  <div key={position} className="position-row">
                    <span>{position}</span>
                    <strong>{count}</strong>
                  </div>
                ))}
              </div>
            </aside>

            <section className="teams-area">
              {teams.length === 0 ? (
                <div className="empty-state">
                  <h2>Todavía no hay equipos generados</h2>
                  <p>Elegí una o más categorías y una formación. Después tocá “Generar equipos”.</p>
                </div>
              ) : (
                <>
                  <section className="panel export-panel">
                    <div>
                      <h2>{teams.length} equipos generados</h2>
                      <p>{completeTeamsCount} completos · {partialTeamsCount} incompletos · {usedPlayers.size} jugadores ubicados · {manualPlayersCount} agregados manualmente · {notUsedPlayers.length} fuera de los 11.</p>
                    </div>
                    <div className="export-actions">
                      <button className="secondary-button" onClick={() => handleExport("png")} disabled={isExporting}>{isExporting ? "Exportando..." : "Exportar todas las imágenes"}</button>
                      <button className="secondary-button" onClick={() => handleExport("pdf")} disabled={isExporting}>{isExporting ? "Exportando..." : "Exportar PDF completo"}</button>
                    </div>
                  </section>

                  <div id="export-area" className="export-area">
                    <div className="export-title">
                      <div>
                        <p className="eyebrow dark-eyebrow">RACING DE MONTEVIDEO SAD</p>
                        <h2>Equipos categorías {selectedCategoryLabel} · {formation}</h2>
                      </div>
                      <strong>{teams.length} equipos</strong>
                    </div>
                    {teams.map((team, teamIndex) => (
                      <TeamCard
                        key={team.name}
                        team={team}
                        teamIndex={teamIndex}
                        teams={teams}
                        categoryLabel={selectedCategoryLabel}
                        isExporting={isExporting}
                        onSwapPlayers={handleSwapPlayers}
                        onInternalSwapPlayers={handleInternalSwapPlayers}
                        onManualNameChange={handleManualNameChange}
                        onRemovePlayer={handleRemovePlayer}
                        onExportTeam={handleExport}
                      />
                    ))}
                  </div>
                </>
              )}
            </section>
          </section>

          {teams.length > 0 && (
            <section className="panel table-panel">
              <h2>Jugadores sin ubicar o fuera de los 11</h2>
              <p className="small-text">Son jugadores válidos de las categorías seleccionadas que no quedaron actualmente en ningún equipo.</p>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Jugador</th>
                      <th>Posición</th>
                      <th>Pierna</th>
                      <th>Fecha nac.</th>
                      <th>Club</th>
                    </tr>
                  </thead>
                  <tbody>
                    {notUsedPlayers.length === 0 ? (
                      <tr>
                        <td colSpan="5">No quedaron jugadores fuera de los 11.</td>
                      </tr>
                    ) : (
                      notUsedPlayers.map((player) => (
                        <tr key={player.id}>
                          <td>{player.name}</td>
                          <td>{player.positionText}</td>
                          <td>{player.leg}</td>
                          <td>{player.birthDateText}</td>
                          <td>{player.club}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {removedPlayers.length > 0 && (
            <section className="panel table-panel">
              <h2>Jugadores quitados manualmente</h2>
              <p className="small-text">Registro de jugadores eliminados de un puesto después del armado automático.</p>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Jugador</th>
                      <th>Origen</th>
                      <th>Salió de</th>
                      <th>Equipo</th>
                      <th>Posición original</th>
                    </tr>
                  </thead>
                  <tbody>
                    {removedPlayers.map((player) => (
                      <tr key={player.id}>
                        <td>{player.name}</td>
                        <td>{player.origin}</td>
                        <td>{player.roleLabel}</td>
                        <td>{player.teamName}</td>
                        <td>{player.positionText}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {(duplicateSummary.samePlayer.length > 0 || duplicateSummary.sameNameDifferentDate.length > 0) && (
            <section className="panel table-panel">
              <h2>Duplicados y nombres repetidos</h2>
              <p className="small-text">Nombre + fecha de nacimiento iguales se consideran el mismo jugador. Mismo nombre con fecha diferente se mantiene como jugadores distintos.</p>

              {duplicateSummary.samePlayer.length > 0 && (
                <div className="duplicate-block">
                  <h3>Repetidos omitidos por coincidir nombre y fecha</h3>
                  <div className="duplicate-list">
                    {duplicateSummary.samePlayer.map((item) => (
                      <div className="duplicate-item" key={`${item.name}-${item.birthDateText}`}>
                        <strong>{item.name}</strong>
                        <span>{item.positionText} · {item.birthDateText} · aparece {item.count} veces · filas {item.rows.join(", ")}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {duplicateSummary.sameNameDifferentDate.length > 0 && (
                <div className="duplicate-block">
                  <h3>Mismo nombre con distinta fecha</h3>
                  <div className="duplicate-list">
                    {duplicateSummary.sameNameDifferentDate.map((item) => (
                      <div className="duplicate-item" key={item.name}>
                        <strong>{item.name}</strong>
                        <span>Aparece como {item.count} jugadores distintos porque la fecha de nacimiento no coincide.</span>
                        <ul>
                          {item.entries.map((entry) => (
                            <li key={`${entry.rowNumber}-${entry.birthDateText}`}>{entry.positionText} · {entry.birthDateText} · fila {entry.rowNumber}</li>
                          ))}
                        </ul>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </section>
          )}

          {invalidPlayers.length > 0 && (
            <section className="panel table-panel">
              <h2>Jugadores con datos a revisar</h2>
              <p className="small-text">No entran en el armado automático hasta corregir fecha de nacimiento o posición.</p>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Jugador</th>
                      <th>Posición</th>
                      <th>Fecha detectada</th>
                      <th>Motivo</th>
                    </tr>
                  </thead>
                  <tbody>
                    {invalidPlayers.map((player) => (
                      <tr key={player.id}>
                        <td>{player.name}</td>
                        <td>{player.positionText || "Sin posición"}</td>
                        <td>{player.birthDateText || "Sin fecha"}</td>
                        <td>{player.category === "Fecha inválida" ? "Fecha inválida" : "Falta nombre o posición"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </>
      )}
    </main>
  );
}

function TeamCard({ team, teamIndex, teams, categoryLabel, isExporting, onSwapPlayers, onInternalSwapPlayers, onManualNameChange, onRemovePlayer, onExportTeam }) {
  const missing = team.picked.filter((item) => !item.player).length;

  return (
    <article className="team-card" data-export-team="true">
      <div className="team-header">
        <div>
          <h2>{team.name}</h2>
          <span>{team.formationName} · Categorías {categoryLabel}</span>
        </div>
        <div className="export-actions">
          <strong className={missing ? "missing-pill" : "ok-pill"}>{missing ? `${missing} faltantes` : "Completo"}</strong>
          {!isExporting && (
            <>
              <button className="secondary-button" type="button" onClick={() => onExportTeam("png", teamIndex)}>Descargar imagen</button>
              <button className="secondary-button" type="button" onClick={() => onExportTeam("pdf", teamIndex)}>Descargar PDF</button>
            </>
          )}
        </div>
      </div>

      <div className="pitch">
        {groupTeam(team).map((group) => (
          <div className={`pitch-line line-${sanitizeFileName(group.line)}`} key={group.line}>
            {getVisualLineItems(group.items).map((item) => {
              const swapOptions = getSwapOptions(teams, teamIndex, item);
              const internalSwapOptions = getInternalSwapOptions(team, item.roleIndex);
              const playerDescription = getPlayerDescription(item);

              return (
                <div className={`player-chip ${!item.player ? "missing" : ""}`} key={`${team.name}-${item.roleIndex}`}>
                  <span>{item.role.label}</span>
                  <strong>{item.player ? item.player.name : "Sin jugador"}</strong>

                  {item.player && <small>{playerDescription}</small>}

                  {(!item.player || item.player.isManual) && !isExporting && (
                    <input
                      type="text"
                      value={item.player?.isManual ? item.player.name : ""}
                      onChange={(event) => onManualNameChange(teamIndex, item.roleIndex, event.target.value)}
                      placeholder="Escribir jugador manual"
                      style={{
                        width: "100%",
                        marginTop: "8px",
                        padding: "8px 10px",
                        borderRadius: "10px",
                        border: "1px solid rgba(0,0,0,0.18)",
                        fontSize: "12px"
                      }}
                    />
                  )}

                  {item.player && !isExporting && (
                    <button
                      type="button"
                      className="secondary-button"
                      onClick={() => onRemovePlayer(teamIndex, item.roleIndex)}
                      style={{
                        width: "100%",
                        marginTop: "8px",
                        padding: "8px 10px",
                        fontSize: "12px"
                      }}
                    >
                      Quitar del 11
                    </button>
                  )}

                  {item.player && !isExporting && internalSwapOptions.length > 0 && (
                    <select
                      className="swap-select"
                      value=""
                      onChange={(event) => {
                        onInternalSwapPlayers(teamIndex, item.roleIndex, event.target.value);
                        event.target.value = "";
                      }}
                    >
                      <option value="">Intercambiar dentro del 11...</option>
                      {internalSwapOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.player.name} · está como {option.role.label} · {option.player.positionText || "Sin posición"}
                        </option>
                      ))}
                    </select>
                  )}

                  {item.player && !isExporting && swapOptions.length > 0 && (
                    <select
                      className="swap-select"
                      value=""
                      onChange={(event) => {
                        onSwapPlayers(teamIndex, item.roleIndex, event.target.value);
                        event.target.value = "";
                      }}
                    >
                      <option value="">Cambiar por otro equipo...</option>
                      {swapOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.player.name} · está en {option.teamName} · {option.player.positionText || "Sin posición"}
                        </option>
                      ))}
                    </select>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </article>
  );
}

export default App;