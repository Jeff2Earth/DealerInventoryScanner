
function normalizeLegacyRow(r) {
  const model = (r.md ?? "").toString().trim();
  return {
    stock: r.s ?? "",
    year: r.y ?? "",
    make: (r.mk ?? "").toString().trim(),
    model,
    type: classifyType(model),
    desc: r.de ?? "",
    status: (r.st ?? "").toString().toUpperCase().trim(),
    color: r.c ?? "",
    odometer: parseNum(r.o),
    vin: (r.v ?? "").toString().trim().toUpperCase(),
    days: parseNum(r.d),
    price: parseMoney(r.p),
    certified: !!r.ce,
    recall: "",
    drivetrain: "",
  };
}

// "2024 Toyota Corolla LE" -> { year: 2024, make: "Toyota", model: "Corolla LE" }
function parseVehicleString(vehicle) {
  const s = (vehicle || "").toString().trim();
  const m = s.match(/^(\d{4})\s+(\S+)\s+(.*)$/);
  if (!m) return { year: null, make: "", model: s };
  return { year: parseInt(m[1], 10), make: m[2], model: m[3].trim() };
}

function normalizeKey(k) {
  return k.replace(/\s+/g, " ").trim().toLowerCase();
}

function getField(row, name) {
  for (const k of Object.keys(row)) {
    if (normalizeKey(k) === name) return row[k];
  }
  return undefined;
}

// Native pricing-export schema: Photos, Autowriter Description, Vehicle,
// Stock #, VIN, Class, Certified, Deleted Date, Status, Recall Status, Body,
// Color, Disp, Price / % Mkt, Last $ Change, Odometer.
//
// Ignored on purpose: Photos, Autowriter Description, Deleted Date, Status,
// Disp — Status in particular is always blank in this export (verified
// against a real pull), so it can't drive the status filter/badge the way
// the legacy CSV format could. Certified and Days are kept: Certified is
// read straight from the export, and Days is left null since this export
