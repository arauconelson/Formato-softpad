/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface DocRow {
  id: string;
  proveedor: string;
  factura: string;
  fechaFactura: string;
  item: string;
  partida: string;
  marca: string;
  modelo: string;
  cantidadUnidadComercial: string;
  tipoUnidadComercial: string;
  cantidadBultos: string;
  claseDeBultos: string;
  precioItemUSD: string;
  paisOrigen: string;
  estado: string;
  pesoNeto: string;
  numeroParte: string;
  descripcion: string;
  caracteristica1: string;
  caracteristica2: string;
  caracteristica3: string;
  caracteristica4: string;
  confidenceMap?: Record<string, 'high' | 'medium' | 'low'>;
}

export type PartialDocRow = Partial<DocRow>;
