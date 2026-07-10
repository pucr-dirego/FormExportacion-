import type {
  ExportExcelPayload,
  ExportFlowAttachment,
  ExportFlowPayload,
  ExportFormState,
} from "../types/exportacion.types";

type BuildExportacionExcelPayloadParams = {
  form: ExportFormState;
  realizadoPor: string;
};

type BuildExportacionFlowPayloadParams = {
  form: ExportFormState;
  realizadoPor: string;
  usuarioCorreo?: string;
};

export function buildExportacionExcelPayload({
  form,
  realizadoPor,
}: BuildExportacionExcelPayloadParams): ExportExcelPayload {
  return {
    "Row ID": "",
    "ID TICKET": "",
    "FOLIO": "",

    "ICOTERM": form.icoterm,
    "TIPO": form.tipo,
    "CANTIDAD CONTENEDORES": form.cantidadContenedores,
    "TAMAÑO/TIPO CONTENEDOR": form.tamanoTipoContenedor,
    "CANTIDAD CONTENEDORES 20": form.cantidadContenedores20,
    "CANTIDAD CONTENEDORES 40": form.cantidadContenedores40,
    "FECHA MATERIAL LISTO": form.fechaMaterialListo,
    "CONSIGNATARIO": form.consignatario.join(", "),
    "ORDEN DE COMPRA SISTEMA": form.ordenCompraSistema,
    "ORDEN DE COMPRA PROVEEDOR": form.ordenCompraProveedor,
    "PUERTO DE SALIDA": form.puertoSalida,
    "PUERTO DE LLEGADA": form.puertoLlegada,
    "PROVEEDOR": form.proveedor,
    "DATOS FISCALES PROVEEDOR": form.datosFiscalesProveedor,
    "INFORMACIÓN CONTACTO PROVEEDOR": form.informacionContactoProveedor,
    "COMENTARIOS": form.comentarios,
    "RELIZADO POR": realizadoPor,

    "ESTADO CORREO": "",
    "FECHA ENVIO CORREO": "",
    "CORREO DESTINO": "",
    "ESTADO COTIZACIÓN": "",
    "REFERENCIA COTIZACIÓN": "",
    "CORREO DEL REALIZADOR": "",
    "ESTADO CORREO  REALIZADOR": "",
    "HORA REGISTRO": "",
    "HORA REGISTRO EN DATE": "",
    "HORA MINIMA DE REGISTRO": "",
    "TIEMPO DE RETRASO": "",
  };
}

export function buildExportacionFlowPayload({
  form,
  realizadoPor,
  usuarioCorreo = "",
}: BuildExportacionFlowPayloadParams): ExportFlowPayload {
  return {
    icoterm: form.icoterm,
    tipo: form.tipo,
    cantidadContenedores: form.cantidadContenedores,
    tamanoTipoContenedor: form.tamanoTipoContenedor,
    cantidadContenedores20: form.cantidadContenedores20,
    cantidadContenedores40: form.cantidadContenedores40,
    fechaMaterialListo: form.fechaMaterialListo,
    consignatario: form.consignatario.join(", "),
    ordenCompraSistema: form.ordenCompraSistema,
    ordenCompraProveedor: form.ordenCompraProveedor,
    puertoSalida: form.puertoSalida,
    puertoLlegada: form.puertoLlegada,
    proveedor: form.proveedor,
    datosFiscalesProveedor: form.datosFiscalesProveedor,
    informacionContacto: form.informacionContactoProveedor,
    informacionContactoProveedor: form.informacionContactoProveedor,
    comentarios: form.comentarios,
    usuarioNombre: realizadoPor,
    usuarioCorreo,
  };
}

export function fileToBase64Clean(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const result = reader.result;

      if (typeof result !== "string") {
        reject(new Error(`No se pudo leer el archivo: ${file.name}`));
        return;
      }

      const base64 = result.includes(",") ? result.split(",")[1] : result;
      resolve(base64);
    };

    reader.onerror = () => {
      reject(new Error(`No se pudo leer el archivo: ${file.name}`));
    };

    reader.readAsDataURL(file);
  });
}

export async function buildExportacionFlowAttachments(
  files: File[]
): Promise<ExportFlowAttachment[]> {
  return Promise.all(
    files.map(async (file) => ({
      name: file.name,
      contentType: file.type || null,
      size: file.size,
      contentBytes: await fileToBase64Clean(file),
    }))
  );
}
