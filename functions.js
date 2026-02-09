/**
 * Halla el valor de venta.
 * Formula: NV = Int(NOVO_COSTO / NOVO_MARGEN) + 1
 * @customfunction
 * @param {number} novo_costo NOVO_COSTO
 * @param {number} novo_margen NOVO_MARGEN (se usará como 1 - novo_margen)
 * @returns {number} El valor de venta calculado.
 */
function nv(novo_costo, novo_margen) {
  // NOVO_MARGEN = 100% - A2
  var real_margen = 1 - novo_margen;

  if (real_margen === 0) {
    return 0; // Evitar división por cero
  }

  // NV = NOVO_COSTO / (1 - NOVO_MARGEN)
  var division = novo_costo / real_margen;

  // NV = Int(NV) + 1
  return Math.floor(division) + 1;
}

// Register functions
CustomFunctions.associate("NV", nv);
