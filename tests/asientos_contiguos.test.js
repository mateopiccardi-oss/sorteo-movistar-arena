const test = require("node:test");
const assert = require("node:assert/strict");
const { loadGs } = require("./_load_gs");

const { _parsearAsientoPdf, _asignarPdfsContiguos } = loadGs();

// Nombre real del sistema de ticketing (Fatboy Slim 05-09-2026)
const mk = (ticket, sector, fila, asiento) =>
  `${ticket} - 05-09-2026 - FATBOY SLIM - ${sector} - ${fila} - ${asiento} - Empleados Movistar Arena  .pdf`;

// Layout real de los 112 PDFs ya enviados + 36 sintéticos con el mismo formato (148 en total).
function layoutFatboy() {
  const filas = [
    ["112", "R", 1, 13], ["112", "P", 1, 11], ["111", "R", 5, 12], ["111", "P", 1, 12],
    ["111", "N", 1, 12], ["111", "J", 4, 12], ["109", "S", 1, 12], ["109", "Q", 1, 12],
    ["109", "M", 1, 12], ["109", "J", 5, 9], ["116", "M", 1, 12], ["116", "K", 1, 13],
    ["116", "L", 1, 12], ["115", "A", 1, 5],
  ];
  const out = [];
  let ticket = 18862745;
  filas.forEach(([sector, fila, desde, hasta]) => {
    for (let a = desde; a <= hasta; a++) out.push(mk(ticket++, sector, fila, a));
  });
  return out;
}

const key = (n) => { const p = _parsearAsientoPdf(n); return p && p.sector + "|" + p.fila; };
const esContiguo = (nombres) => {
  const ps = nombres.map(_parsearAsientoPdf);
  if (ps.some((p) => !p)) return false;
  const k = ps[0].sector + "|" + ps[0].fila;
  if (ps.some((p) => p.sector + "|" + p.fila !== k)) return false;
  const as = ps.map((p) => p.asiento).sort((a, b) => a - b);
  return as.every((a, i) => i === 0 || a === as[i - 1] + 1);
};

test("parsea sector, fila y asiento de un nombre real", () => {
  const p = _parsearAsientoPdf("18862745 - 05-09-2026 - FATBOY SLIM - 112 - R - 13 - Empleados Movistar Arena  .pdf");
  assert.deepEqual(p, { sector: "112", fila: "R", asiento: 13 });
});

test("devuelve null si el nombre no tiene el formato del ticketing", () => {
  assert.equal(_parsearAsientoPdf("entrada_1.pdf"), null);
  assert.equal(_parsearAsientoPdf("18862745 - FATBOY SLIM - 112 - R - trece - x.pdf"), null);
});

test("cada ganador recibe asientos consecutivos de la misma fila cuando hay corrida disponible", () => {
  const pdfs = layoutFatboy();
  const cantidades = Array(68).fill(2); // 68 ganadores x 2 = 136 <= 148
  const r = _asignarPdfsContiguos(pdfs, cantidades);
  assert.equal(r.asignacion.length, 68);
  r.asignacion.forEach((nombres, i) => {
    assert.equal(nombres.length, 2, "ganador " + i + " recibe 2");
    if (!r.noContiguos.includes(i)) assert.ok(esContiguo(nombres), "ganador " + i + " no contiguo: " + nombres.join(" | "));
  });
  // Filas impares: 112 R (13), 111 J (9), 109 J (5), 116 K (13), 115 A (5) → 5 sueltos → como mucho 3 ganadores no contiguos
  assert.ok(r.noContiguos.length <= 3, "noContiguos=" + r.noContiguos.length);
});

test("no asigna dos veces el mismo PDF y respeta la cantidad de cada ganador", () => {
  const pdfs = layoutFatboy();
  const cantidades = [2, 2, 4, 2, 6, 2, 4, 2];
  const r = _asignarPdfsContiguos(pdfs, cantidades);
  const todos = r.asignacion.flat();
  assert.equal(new Set(todos).size, todos.length);
  r.asignacion.forEach((n, i) => assert.equal(n.length, cantidades[i]));
});

test("los grupos grandes se ubican primero aunque vengan últimos en la lista", () => {
  // Solo una fila de 6 disponible: si se reparte en orden, el de 6 no entra contiguo.
  const pdfs = [1, 2, 3, 4, 5, 6].map((a) => mk(100 + a, "116", "M", a))
    .concat([1, 2, 3, 4].map((a) => mk(200 + a, "115", "A", a)))
    .concat([7, 8].map((a) => mk(300 + a, "109", "J", a)));
  const r = _asignarPdfsContiguos(pdfs, [2, 2, 6]);
  assert.ok(esContiguo(r.asignacion[2]), "el de 6 debe ir contiguo: " + r.asignacion[2].join(" | "));
  assert.ok(esContiguo(r.asignacion[0]));
  assert.ok(esContiguo(r.asignacion[1]));
  assert.deepEqual(r.noContiguos, []);
});

test("sin corrida disponible asigna igual y lo informa en noContiguos", () => {
  // 3 filas con 1 asiento suelto cada una, un ganador de 2
  const pdfs = [mk(1, "112", "R", 13), mk(2, "111", "J", 4), mk(3, "109", "J", 9)];
  const r = _asignarPdfsContiguos(pdfs, [2]);
  assert.equal(r.asignacion[0].length, 2);
  assert.deepEqual(r.noContiguos, [0]);
});

test("un PDF con nombre no parseable se usa solo al final, como suelto", () => {
  const pdfs = ["entrada_1.pdf", mk(1, "112", "R", 1), mk(2, "112", "R", 2), mk(3, "111", "P", 1)];
  const r = _asignarPdfsContiguos(pdfs, [2, 2]);
  assert.deepEqual(r.asignacion[0].map(key), ["112|R", "112|R"]);
  assert.ok(r.asignacion[1].includes("entrada_1.pdf"));
  assert.deepEqual(r.noContiguos, [1]);
});

test("a igual cantidad, el primero de la lista recibe la primera corrida (orden de la planilla)", () => {
  const pdfs = [mk(1, "109", "M", 1), mk(2, "109", "M", 2), mk(3, "116", "K", 1), mk(4, "116", "K", 2)];
  const r = _asignarPdfsContiguos(pdfs, [2, 2]);
  assert.equal(key(r.asignacion[0][0]), "109|M");
  assert.equal(key(r.asignacion[1][0]), "116|K");
});
