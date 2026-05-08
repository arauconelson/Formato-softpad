/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleGenAI, Type } from "@google/genai";
import { DocRow } from "./types";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

const docRowSchema = {
  type: Type.ARRAY,
  items: {
    type: Type.OBJECT,
    properties: {
      proveedor: { type: Type.STRING },
      factura: { type: Type.STRING },
      fechaFactura: { type: Type.STRING },
      item: { type: Type.STRING },
      partida: { type: Type.STRING },
      marca: { type: Type.STRING },
      modelo: { type: Type.STRING },
      cantidadUnidadComercial: { type: Type.STRING },
      tipoUnidadComercial: { type: Type.STRING },
      cantidadBultos: { type: Type.STRING },
      claseDeBultos: { type: Type.STRING },
      precioItemUSD: { type: Type.STRING },
      paisOrigen: { type: Type.STRING },
      estado: { type: Type.STRING },
      pesoNeto: { type: Type.STRING },
      numeroParte: { type: Type.STRING },
      descripcion: { type: Type.STRING },
      caracteristica1: { type: Type.STRING },
      caracteristica2: { type: Type.STRING },
      caracteristica3: { type: Type.STRING },
      caracteristica4: { type: Type.STRING },
      confidenceMap: {
        type: Type.OBJECT,
        properties: {
          proveedor: { type: Type.STRING, enum: ["high", "medium", "low"] },
          factura: { type: Type.STRING, enum: ["high", "medium", "low"] },
          fechaFactura: { type: Type.STRING, enum: ["high", "medium", "low"] },
          item: { type: Type.STRING, enum: ["high", "medium", "low"] },
          partida: { type: Type.STRING, enum: ["high", "medium", "low"] },
          marca: { type: Type.STRING, enum: ["high", "medium", "low"] },
          modelo: { type: Type.STRING, enum: ["high", "medium", "low"] },
          cantidadUnidadComercial: { type: Type.STRING, enum: ["high", "medium", "low"] },
          tipoUnidadComercial: { type: Type.STRING, enum: ["high", "medium", "low"] },
          cantidadBultos: { type: Type.STRING, enum: ["high", "medium", "low"] },
          claseDeBultos: { type: Type.STRING, enum: ["high", "medium", "low"] },
          precioItemUSD: { type: Type.STRING, enum: ["high", "medium", "low"] },
          paisOrigen: { type: Type.STRING, enum: ["high", "medium", "low"] },
          estado: { type: Type.STRING, enum: ["high", "medium", "low"] },
          pesoNeto: { type: Type.STRING, enum: ["high", "medium", "low"] },
          numeroParte: { type: Type.STRING, enum: ["high", "medium", "low"] },
          descripcion: { type: Type.STRING, enum: ["high", "medium-low"] },
          caracteristica1: { type: Type.STRING, enum: ["high", "medium", "low"] },
          caracteristica2: { type: Type.STRING, enum: ["high", "medium", "low"] },
          caracteristica3: { type: Type.STRING, enum: ["high", "medium", "low"] },
          caracteristica4: { type: Type.STRING, enum: ["high", "medium", "low"] },
        }
      }
    },
    required: ["descripcion", "confidenceMap"],
  },
};

export interface FileData {
  base64: string;
  mimeType: string;
  name?: string;
}

export async function extractTableFromImage(
  mainFile: FileData, 
  complementaryFiles: FileData[] = []
): Promise<DocRow[]> {
  try {
    const contents = {
        parts: [
          {
            text: `Eres un experto Senior en Clasificación Arancelaria, Aduanas y Logística internacional con acceso a múltiples documentos.
            
            JERARQUÍA DE DOCUMENTOS:
            1. DOCUMENTO PRINCIPAL: Es el primer archivo adjunto (la Factura o Póliza). Tu tarea primaria es extraer los items de este documento.
            2. DOCUMENTOS DE CONTEXTO: Son los archivos adicionales. Úsalos como REFERENCIA TÉCNICA para resolver dudas, especialmente sobre la PARTIDA ARANCELARIA.
            
            INSTRUCCIÓN DE ANÁLISIS:
            - Procesa todos los documentos adjuntos, analizando todas las páginas u hojas presentes.
            - Busca en los documentos de contexto descripciones técnicas que coincidan con los productos de la factura para asignar la partida arancelaria exacta según el ARANCEL DE ADUANAS 2022.
            - Si un dato en la factura es impreciso pero el documento de contexto lo aclara, usa la información del contexto y marca confianza como "high".

            REGLAS DE NEGOCIO CRÍTICAS:
            1. ESTADO: Coloca SIEMPRE el número "10" (representa nuevo/estándar). Es obligatorio.
            2. PARTIDA ARANCELARIA: Formato xxxx.xx.xx.xx (10 dígitos). 
               - Ejemplo Smartphones: 8517.13.00.00.
               - Ejemplo Otros móviles: 8517.14.00.00.
            3. PAIS ORIGEN: Código ISO de 2 letras.
            4. FECHA DE FACTURA: Formato DD/MM/AAAA.
            5. MODELO: Si dice "Remark", es el Modelo.
            6. TIPO UNIDAD COMERCIAL: "U".
            7. CLASE DE BULTOS: "BUL".
            8. NUMERO PARTE: Vacío por defecto.

            EVALUACIÓN DE CONFIANZA (confidenceMap) - REGLAS ESTRICTAS:
            - "high" (Verde): Datos 100% legibles y Partida Arancelaria confirmada.
            - "medium" (Amarillo): USAR ante cualquier ambigüedad en la partida, descripciones genéricas o falta de sustento técnico en el contexto. SIEMPRE prioriza "medium" ante la mínima duda.
            - "low" (Rojo): Dato ilegible o no disponible en ningún documento.
            
            Extrae del DOCUMENTO PRINCIPAL: PROVEEDOR, FACTURA, FECHA, ITEM, PARTIDA, MARCA, MODELO, CANTIDADES, PRECIOS, PAIS, PESO y DESCRIPCIÓN con sus CARACTERISTICAS.`,
          },
          // Archivo principal (Factura)
          {
            inlineData: {
              data: mainFile.base64.split(",")[1] || mainFile.base64,
              mimeType: mainFile.mimeType,
            },
          },
          // Archivos complementarios (Contexto)
          ...complementaryFiles.map(file => ({
            inlineData: {
              data: file.base64.split(",")[1] || file.base64,
              mimeType: file.mimeType,
            }
          }))
        ],
    };

    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: contents,
      config: {
        responseMimeType: "application/json",
        responseSchema: docRowSchema,
      },
    });

    const rows = JSON.parse(response.text || "[]") as any[];
    return rows.map((row, index) => ({
      ...row,
      id: crypto.randomUUID(),
    })) as DocRow[];
  } catch (error) {
    console.error("Error extracting data:", error);
    throw error;
  }
}
