export type ExportFormState = {
  icoterm: string;
  tipo: string;
  cantidadContenedores: string;
  tamanoTipoContenedor: string;
  cantidadContenedores20: string;
  cantidadContenedores40: string;
  fechaMaterialListo: string;
  consignatario: string[];
  ordenCompraSistema: string;
  ordenCompraProveedor: string;
  puertoSalida: string;
  puertoLlegada: string;
  proveedor: string;
  datosFiscalesProveedor: string;
  informacionContactoProveedor: string;
  comentarios: string;
};

export type ExportFormErrors = Partial<Record<keyof ExportFormState, string>>;

export type ExportSubmitMessage = {
  type: "success" | "error";
  text: string;
};

export type ExportExcelPayload = {
  "Row ID": string;
  "ID TICKET": string;
  "FOLIO": string;

  "ICOTERM": string;
  "TIPO": string;
  "CANTIDAD CONTENEDORES": string;
  "TAMAÑO/TIPO CONTENEDOR": string;
  "CANTIDAD CONTENEDORES 20": string;
  "CANTIDAD CONTENEDORES 40": string;
  "FECHA MATERIAL LISTO": string;
  "CONSIGNATARIO": string;
  "ORDEN DE COMPRA SISTEMA": string;
  "ORDEN DE COMPRA PROVEEDOR": string;
  "PUERTO DE SALIDA": string;
  "PUERTO DE LLEGADA": string;
  "PROVEEDOR": string;
  "DATOS FISCALES PROVEEDOR": string;
  "INFORMACIÓN CONTACTO PROVEEDOR": string;
  "COMENTARIOS": string;
  "RELIZADO POR": string;

  "ESTADO CORREO": string;
  "FECHA ENVIO CORREO": string;
  "CORREO DESTINO": string;
  "ESTADO COTIZACIÓN": string;
  "REFERENCIA COTIZACIÓN": string;
  "CORREO DEL REALIZADOR": string;
  "ESTADO CORREO  REALIZADOR": string;
  "HORA REGISTRO": string;
  "HORA REGISTRO EN DATE": string;
  "HORA MINIMA DE REGISTRO": string;
  "TIEMPO DE RETRASO": string;
};

export type ExportFlowPayload = {
  icoterm: string;
  tipo: string;
  cantidadContenedores: string;
  tamanoTipoContenedor: string;
  cantidadContenedores20: string;
  cantidadContenedores40: string;
  fechaMaterialListo: string;
  consignatario: string;
  ordenCompraSistema: string;
  ordenCompraProveedor: string;
  puertoSalida: string;
  puertoLlegada: string;
  proveedor: string;
  datosFiscalesProveedor: string;
  informacionContacto: string;
  informacionContactoProveedor: string;
  comentarios: string;
  usuarioNombre: string;
  usuarioCorreo: string;
};

export type ExportFlowAttachment = {
  name: string;
  contentBytes: string;
  contentType: string | null;
  size: number;
};
