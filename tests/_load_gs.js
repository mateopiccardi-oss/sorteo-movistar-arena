// Carga sorteo_script.gs en un contexto Node aislado y expone sus funciones puras.
// Los resultados se clonan vía JSON para que tengan los prototipos del contexto del test
// (deepEqual estricto compara prototipos y los del vm son distintos).
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const EXPORTS = ["_parsearAsientoPdf", "_asignarPdfsContiguos"];

function loadGs() {
  const src = fs.readFileSync(path.join(__dirname, "..", "sorteo_script.gs"), "utf8");
  const ctx = { console, Logger: { log() {} } };
  vm.createContext(ctx);
  const tail = "\n;this.__exports={" + EXPORTS.map((n) => n + ":typeof " + n + "==='function'?" + n + ":undefined").join(",") + "};";
  vm.runInContext(src + tail, ctx, { filename: "sorteo_script.gs" });
  const out = {};
  EXPORTS.forEach((n) => {
    const fn = ctx.__exports[n];
    out[n] = fn ? (...args) => JSON.parse(JSON.stringify(fn(...args))) : undefined;
  });
  return out;
}
module.exports = { loadGs };
