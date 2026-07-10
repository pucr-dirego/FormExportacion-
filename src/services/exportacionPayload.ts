import type {
  ExportExcelPayload,
  ExportFormState,
} from "../types/exportacion.types";

type BuildExportacionExcelPayloadParams = {
  form: ExportFormState;
  realizadoPor: string;
};

export function buildExportacionExcelPayload({
  form,
  realizadoPor,
}: BuildExportacionExcelPayloadParams): ExportExcelPayload {
  return {
    /**
     * Columnas que existen en Excel,
     * pero NO se muestran en el formulario.
     * Se mandan vacías.
     */
    "Row ID": "",
    "ID TICKET": "",
    "FOLIO": "",

    /**
     * Columnas capturadas desde el formulario.
     * Orden visible: D a T.
     */
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

    /**
     * Columnas operativas/manuales.
     * Existen en Excel, pero se guardan vacías.
     */
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