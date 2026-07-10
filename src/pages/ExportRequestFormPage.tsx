import { useMemo, useState, type ChangeEvent, type FormEvent } from "react";
import "./ExportRequestFormPage.css";

import {
  cantidadContenedoresOptions,
  consignatarioOptions,
  incotermOptions,
  initialExportFormState,
  proveedorOptions,
  puertoLlegadaOptions,
  puertoSalidaOptions,
  tamanoContenedorOptions,
  tipoOptions,
} from "../constants/exportacionForm.constants";

import {
  buildExportacionFlowAttachments,
  buildExportacionFlowPayload,
} from "../services/exportacionPayload";

import { TEST_RequerimientosExportaci_nService } from "../generated/services/TEST_RequerimientosExportaci_nService";

import type {
  ExportFormErrors,
  ExportFormState,
  ExportSubmitMessage,
} from "../types/exportacion.types";

type TextExportFormField = Exclude<keyof ExportFormState, "consignatario">;

function getCurrentUserName() {
  /**
   * Temporal:
   * Después reemplazamos esto con el usuario real de Power Apps/Teams.
   */
  return "Usuario conectado";
}

const EXPORT_EMAIL_RECIPIENT = "victor.mendoza@dirego.com";
const MAX_ATTACHMENT_FILES = 5;
const MAX_ATTACHMENT_SIZE_BYTES = 10 * 1024 * 1024;
const ACCEPTED_ATTACHMENT_TYPES = ".pdf,.xls,.xlsx";
const ALLOWED_ATTACHMENT_EXTENSIONS = ["pdf", "xls", "xlsx"];

function formatFileSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;

  const kilobytes = bytes / 1024;
  if (kilobytes < 1024) return `${kilobytes.toFixed(1)} KB`;

  const megabytes = kilobytes / 1024;
  return `${megabytes.toFixed(1)} MB`;
}

function getFileExtension(fileName: string) {
  const parts = fileName.split(".");

  if (parts.length <= 1) return "";

  return parts[parts.length - 1].trim().toLowerCase();
}

function isAllowedAttachment(file: File) {
  const extension = getFileExtension(file.name);
  return ALLOWED_ATTACHMENT_EXTENSIONS.includes(extension);
}

export function ExportRequestFormPage() {
  const [form, setForm] = useState<ExportFormState>(initialExportFormState);
  const [errors, setErrors] = useState<ExportFormErrors>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitMessage, setSubmitMessage] = useState<ExportSubmitMessage | null>(null);
  const [isConsignatarioModalOpen, setIsConsignatarioModalOpen] = useState(false);
  const [consignatarioSearch, setConsignatarioSearch] = useState("");
  const [isProveedorSuggestionsOpen, setIsProveedorSuggestionsOpen] = useState(false);
  const [attachedFiles, setAttachedFiles] = useState<File[]>([]);
  const [attachmentError, setAttachmentError] = useState("");

  const realizadoPor = useMemo(() => getCurrentUserName(), []);

  const consignatarioSummary = useMemo(() => {
    if (form.consignatario.length === 0) {
      return "Selecciona consignatario(s)";
    }

    if (form.consignatario.length === 1) {
      return form.consignatario[0];
    }

    return `${form.consignatario.length} consignatarios seleccionados`;
  }, [form.consignatario]);

  const filteredConsignatarioOptions = useMemo(() => {
    const search = consignatarioSearch.trim().toLowerCase();

    if (!search) return consignatarioOptions;

    return consignatarioOptions.filter((option) =>
      option.toLowerCase().includes(search)
    );
  }, [consignatarioSearch]);

  const filteredProveedorOptions = useMemo(() => {
    const search = form.proveedor.trim().toLowerCase();

    if (!search) {
      return proveedorOptions.slice(0, 12);
    }

    const startsWithMatches = proveedorOptions.filter((option) =>
      option.toLowerCase().startsWith(search)
    );

    const containsMatches = proveedorOptions.filter((option) => {
      const normalizedOption = option.toLowerCase();

      return (
        normalizedOption.includes(search) &&
        !normalizedOption.startsWith(search)
      );
    });

    return [...startsWithMatches, ...containsMatches].slice(0, 12);
  }, [form.proveedor]);

  const updateField = (field: TextExportFormField, value: string) => {
    setForm((prev) => ({
      ...prev,
      [field]: value,
    }));

    setErrors((prev) => ({
      ...prev,
      [field]: "",
    }));

    setSubmitMessage(null);
  };

  const handleSelectProveedor = (proveedor: string) => {
    setForm((prev) => ({
      ...prev,
      proveedor,
    }));

    setErrors((prev) => ({
      ...prev,
      proveedor: "",
    }));

    setIsProveedorSuggestionsOpen(false);
    setSubmitMessage(null);
  };

  const handleAttachmentChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFiles: File[] = Array.from(event.target.files ?? []);

    if (selectedFiles.length === 0) return;

    const nextFiles = [...attachedFiles];
    const messages: string[] = [];
    const existingFileKeys = new Set(
      nextFiles.map((file) => `${file.name}-${file.size}-${file.lastModified}`)
    );

    selectedFiles.forEach((file) => {
      const fileKey = `${file.name}-${file.size}-${file.lastModified}`;

      if (nextFiles.length >= MAX_ATTACHMENT_FILES) {
        messages.push(`Solo puedes adjuntar hasta ${MAX_ATTACHMENT_FILES} archivos.`);
        return;
      }

      if (existingFileKeys.has(fileKey)) {
        messages.push(`El archivo ${file.name} ya fue agregado.`);
        return;
      }

      if (!isAllowedAttachment(file)) {
        messages.push(`El archivo ${file.name} no es válido. Solo se permiten PDF, XLS y XLSX.`);
        return;
      }

      if (file.size > MAX_ATTACHMENT_SIZE_BYTES) {
        messages.push(`El archivo ${file.name} supera el límite de 10 MB.`);
        return;
      }

      nextFiles.push(file);
      existingFileKeys.add(fileKey);
    });

    setAttachedFiles(nextFiles);
    setAttachmentError(messages.join(" "));
    setSubmitMessage(null);

    event.target.value = "";
  };

  const handleRemoveAttachment = (indexToRemove: number) => {
    setAttachedFiles((prev) => prev.filter((_, index) => index !== indexToRemove));
    setAttachmentError("");
    setSubmitMessage(null);
  };

  const handleClearAttachments = () => {
    setAttachedFiles([]);
    setAttachmentError("");
    setSubmitMessage(null);
  };

  const toggleConsignatario = (value: string) => {
    setForm((prev) => {
      const alreadySelected = prev.consignatario.includes(value);

      return {
        ...prev,
        consignatario: alreadySelected
          ? prev.consignatario.filter((item) => item !== value)
          : [...prev.consignatario, value],
      };
    });

    setErrors((prev) => ({
      ...prev,
      consignatario: "",
    }));

    setSubmitMessage(null);
  };

  const handleSelectAllConsignatarios = () => {
    setForm((prev) => ({
      ...prev,
      consignatario: consignatarioOptions,
    }));

    setErrors((prev) => ({
      ...prev,
      consignatario: "",
    }));

    setSubmitMessage(null);
  };

  const handleClearConsignatarios = () => {
    setForm((prev) => ({
      ...prev,
      consignatario: [],
    }));

    setErrors((prev) => ({
      ...prev,
      consignatario: "",
    }));

    setSubmitMessage(null);
  };

  const getCantidad20Options = () => {
    const total = Number(form.cantidadContenedores || 0);

    if (total <= 1) return [];

    return Array.from({ length: total - 1 }, (_, index) => String(index + 1));
  };

  const handleCantidadContenedoresChange = (value: string) => {
    setForm((prev) => {
      const total = Number(value || 0);

      if (prev.tamanoTipoContenedor === "20") {
        return {
          ...prev,
          cantidadContenedores: value,
          cantidadContenedores20: value,
          cantidadContenedores40: total > 0 ? "0" : "",
        };
      }

      if (prev.tamanoTipoContenedor === "40") {
        return {
          ...prev,
          cantidadContenedores: value,
          cantidadContenedores20: total > 0 ? "0" : "",
          cantidadContenedores40: value,
        };
      }

      if (prev.tamanoTipoContenedor === "Varios") {
        if (total <= 1) {
          return {
            ...prev,
            cantidadContenedores: value,
            cantidadContenedores20: "",
            cantidadContenedores40: "",
          };
        }

        const current20 = Number(prev.cantidadContenedores20 || 1);
        const safe20 = Math.min(Math.max(current20, 1), total - 1);
        const auto40 = total - safe20;

        return {
          ...prev,
          cantidadContenedores: value,
          cantidadContenedores20: String(safe20),
          cantidadContenedores40: String(auto40),
        };
      }

      return {
        ...prev,
        cantidadContenedores: value,
        cantidadContenedores20: "",
        cantidadContenedores40: "",
      };
    });

    setErrors((prev) => ({
      ...prev,
      cantidadContenedores: "",
      cantidadContenedores20: "",
      cantidadContenedores40: "",
    }));

    setSubmitMessage(null);
  };

  const handleTamanoContenedorChange = (value: string) => {
    setForm((prev) => {
      const total = Number(prev.cantidadContenedores || 0);

      if (value === "20") {
        return {
          ...prev,
          tamanoTipoContenedor: value,
          cantidadContenedores20: total > 0 ? String(total) : "",
          cantidadContenedores40: total > 0 ? "0" : "",
        };
      }

      if (value === "40") {
        return {
          ...prev,
          tamanoTipoContenedor: value,
          cantidadContenedores20: total > 0 ? "0" : "",
          cantidadContenedores40: total > 0 ? String(total) : "",
        };
      }

      if (value === "Varios") {
        if (total <= 1) {
          return {
            ...prev,
            tamanoTipoContenedor: value,
            cantidadContenedores20: "",
            cantidadContenedores40: "",
          };
        }

        return {
          ...prev,
          tamanoTipoContenedor: value,
          cantidadContenedores20: "1",
          cantidadContenedores40: String(total - 1),
        };
      }

      return {
        ...prev,
        tamanoTipoContenedor: value,
        cantidadContenedores20: "",
        cantidadContenedores40: "",
      };
    });

    setErrors((prev) => ({
      ...prev,
      tamanoTipoContenedor: "",
      cantidadContenedores20: "",
      cantidadContenedores40: "",
    }));

    setSubmitMessage(null);
  };

  const handleCantidad20VariosChange = (value: string) => {
    const total = Number(form.cantidadContenedores || 0);
    const cantidad20 = Number(value || 0);
    const cantidad40 = total - cantidad20;

    setForm((prev) => ({
      ...prev,
      cantidadContenedores20: value,
      cantidadContenedores40: String(cantidad40),
    }));

    setErrors((prev) => ({
      ...prev,
      cantidadContenedores20: "",
      cantidadContenedores40: "",
    }));

    setSubmitMessage(null);
  };

  const validateForm = () => {
    const nextErrors: ExportFormErrors = {};

    if (!form.icoterm.trim()) nextErrors.icoterm = "Selecciona el ICOTERM.";
    if (!form.tipo.trim()) nextErrors.tipo = "Selecciona el tipo de operación.";

    if (!form.cantidadContenedores.trim()) {
      nextErrors.cantidadContenedores = "Selecciona la cantidad total de contenedores.";
    }

    if (!form.tamanoTipoContenedor.trim()) {
      nextErrors.tamanoTipoContenedor = "Selecciona el tamaño/tipo de contenedor.";
    }

    if (
      form.tamanoTipoContenedor === "Varios" &&
      Number(form.cantidadContenedores || 0) <= 1
    ) {
      nextErrors.tamanoTipoContenedor =
        "Para usar varios tamaños, selecciona al menos 2 contenedores.";
    }

    if (
      form.tamanoTipoContenedor === "Varios" &&
      !form.cantidadContenedores20.trim()
    ) {
      nextErrors.cantidadContenedores20 = "Selecciona la cantidad de contenedores de 20.";
    }

    if (!form.fechaMaterialListo.trim()) {
      nextErrors.fechaMaterialListo = "Indica la fecha en que el material estará listo.";
    }

    if (form.consignatario.length === 0) {
      nextErrors.consignatario = "Selecciona al menos un consignatario.";
    }

    if (!form.ordenCompraSistema.trim()) {
      nextErrors.ordenCompraSistema = "Captura la orden de compra del sistema.";
    }

    if (!form.ordenCompraProveedor.trim()) {
      nextErrors.ordenCompraProveedor = "Captura la orden de compra del proveedor.";
    }

    if (!form.puertoSalida.trim()) {
      nextErrors.puertoSalida = "Selecciona el puerto de salida.";
    }

    if (!form.puertoLlegada.trim()) {
      nextErrors.puertoLlegada = "Selecciona el puerto de llegada.";
    }

    if (!form.proveedor.trim()) nextErrors.proveedor = "Captura el proveedor.";

    if (!form.datosFiscalesProveedor.trim()) {
      nextErrors.datosFiscalesProveedor = "Captura los datos fiscales del proveedor.";
    }

    if (!form.informacionContactoProveedor.trim()) {
      nextErrors.informacionContactoProveedor =
        "Captura la información de contacto del proveedor.";
    }

    setErrors(nextErrors);

    return Object.keys(nextErrors).length === 0;
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!validateForm()) {
      setSubmitMessage({
        type: "error",
        text: "Revisa los campos marcados antes de enviar el requerimiento.",
      });
      return;
    }

    setIsSubmitting(true);
    setSubmitMessage(null);

    try {
      const payload = buildExportacionFlowPayload({
        form,
        realizadoPor,
        usuarioCorreo: "",
      });

      const attachments = await buildExportacionFlowAttachments(attachedFiles);

      console.log("payloadJson listo para Power Automate:", payload);
      console.log("attachmentsJson listo para Power Automate:", attachments);
      console.log("Correo destino:", EXPORT_EMAIL_RECIPIENT);

      const flowResult = await TEST_RequerimientosExportaci_nService.Run({
        /**
         * Ojo:
         * text_1 = payloadJson
         * text   = attachmentsJson
         *
         * Así lo generó el modelo:
         * // payloadJson
         * text_1: string;
         * // attachmentsJson
         * text: string;
         */
        text_1: JSON.stringify(payload),
        text: JSON.stringify(attachments),
      });

      if (!flowResult.success) {
        throw new Error("El flujo no respondió correctamente.");
      }

      if (flowResult.data?.status !== "ok") {
        throw new Error(
          flowResult.data?.message || "El flujo devolvió una respuesta de error."
        );
      }

      setSubmitMessage({
        type: "success",
        text: flowResult.data?.message || "Requerimiento enviado correctamente.",
      });

      setForm(initialExportFormState);
      setIsConsignatarioModalOpen(false);
      setConsignatarioSearch("");
      setIsProveedorSuggestionsOpen(false);
      setAttachedFiles([]);
      setAttachmentError("");
    } catch (error) {
      console.error(error);

      setSubmitMessage({
        type: "error",
        text:
          error instanceof Error
            ? error.message
            : "No se pudo enviar el requerimiento. Intenta nuevamente.",
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleClear = () => {
    const confirmClear = window.confirm(
      "¿Deseas limpiar el formulario? Se perderán los datos capturados."
    );

    if (!confirmClear) return;

    setForm(initialExportFormState);
    setErrors({});
    setSubmitMessage(null);
    setIsConsignatarioModalOpen(false);
    setConsignatarioSearch("");
    setIsProveedorSuggestionsOpen(false);
    setAttachedFiles([]);
    setAttachmentError("");
  };

  return (
    <main className="export-page">
      <section className="export-hero">
        <div>
          <h1>Requerimiento de exportación</h1>
          <p>
            Captura la información necesaria para iniciar el seguimiento de una operación de exportación.
            Los campos operativos internos permanecerán vacíos en Excel para su llenado manual posterior.
          </p>
        </div>

        <div className="export-user-card">
          <span>Registrado por</span>
          <strong>{realizadoPor}</strong>
          <small>Este dato se llenará automáticamente.</small>
        </div>
      </section>

      <form className="export-form" onSubmit={handleSubmit}>
        <section className="form-section">
          <div className="section-header">
            <span>01</span>
            <div>
              <h2>Información inicial</h2>
              <p>Campos D y E del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span>ICOTERM *</span>
              <select
                value={form.icoterm}
                onChange={(event) => updateField("icoterm", event.target.value)}
              >
                <option value="">Selecciona una opción</option>
                {incotermOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.icoterm && <small>{errors.icoterm}</small>}
            </label>

            <label className="form-field">
              <span>Tipo *</span>
              <select
                value={form.tipo}
                onChange={(event) => updateField("tipo", event.target.value)}
              >
                <option value="">Selecciona una opción</option>
                {tipoOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.tipo && <small>{errors.tipo}</small>}
            </label>
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>02</span>
            <div>
              <h2>Contenedores</h2>
              <p>Campos F, G, H e I del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span>Cantidad contenedores *</span>
              <select
                value={form.cantidadContenedores}
                onChange={(event) => handleCantidadContenedoresChange(event.target.value)}
              >
                <option value="">Selecciona la cantidad</option>
                {cantidadContenedoresOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.cantidadContenedores && <small>{errors.cantidadContenedores}</small>}
            </label>

            <label className="form-field">
              <span>Tamaño/tipo contenedor *</span>
              <select
                value={form.tamanoTipoContenedor}
                onChange={(event) => handleTamanoContenedorChange(event.target.value)}
              >
                <option value="">Selecciona una opción</option>
                {tamanoContenedorOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.tamanoTipoContenedor && <small>{errors.tamanoTipoContenedor}</small>}
            </label>

            {form.tamanoTipoContenedor === "20" && (
              <>
                <label className="form-field">
                  <span>Cantidad contenedores 20</span>
                  <input
                    className="readonly-field"
                    type="text"
                    value={form.cantidadContenedores20}
                    readOnly
                  />
                </label>

                <label className="form-field">
                  <span>Cantidad contenedores 40</span>
                  <input
                    className="readonly-field"
                    type="text"
                    value={form.cantidadContenedores40}
                    readOnly
                  />
                </label>
              </>
            )}

            {form.tamanoTipoContenedor === "40" && (
              <>
                <label className="form-field">
                  <span>Cantidad contenedores 20</span>
                  <input
                    className="readonly-field"
                    type="text"
                    value={form.cantidadContenedores20}
                    readOnly
                  />
                </label>

                <label className="form-field">
                  <span>Cantidad contenedores 40</span>
                  <input
                    className="readonly-field"
                    type="text"
                    value={form.cantidadContenedores40}
                    readOnly
                  />
                </label>
              </>
            )}

            {form.tamanoTipoContenedor === "Varios" && (
              <>
                <label className="form-field">
                  <span>Cantidad de contenedores de 20 *</span>
                  <select
                    value={form.cantidadContenedores20}
                    onChange={(event) => handleCantidad20VariosChange(event.target.value)}
                    disabled={Number(form.cantidadContenedores || 0) <= 1}
                  >
                    <option value="">Selecciona cantidad</option>
                    {getCantidad20Options().map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                  {errors.cantidadContenedores20 && (
                    <small>{errors.cantidadContenedores20}</small>
                  )}
                </label>

                <label className="form-field">
                  <span>Cantidad de contenedores de 40</span>
                  <input
                    className="readonly-field"
                    type="text"
                    value={form.cantidadContenedores40}
                    readOnly
                  />
                </label>
              </>
            )}
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>03</span>
            <div>
              <h2>Fecha y consignatario</h2>
              <p>Campos J y K del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span>Fecha material listo *</span>
              <input
                type="date"
                value={form.fechaMaterialListo}
                onChange={(event) => updateField("fechaMaterialListo", event.target.value)}
              />
              {errors.fechaMaterialListo && <small>{errors.fechaMaterialListo}</small>}
            </label>

            <div className="form-field form-field-wide">
              <span>Consignatario *</span>

              <button
                type="button"
                className={`consignatario-trigger ${
                  form.consignatario.length > 0 ? "consignatario-trigger-filled" : ""
                }`}
                onClick={() => setIsConsignatarioModalOpen(true)}
              >
                <span>{consignatarioSummary}</span>
                <strong>Abrir selector</strong>
              </button>

              {form.consignatario.length > 0 && (
                <div className="selected-tags">
                  {form.consignatario.map((item) => (
                    <span key={item}>{item}</span>
                  ))}
                </div>
              )}

              {errors.consignatario && <small>{errors.consignatario}</small>}
            </div>
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>04</span>
            <div>
              <h2>Órdenes y ruta</h2>
              <p>Campos L, M, N y O del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span>Orden de compra sistema *</span>
              <input
                type="text"
                value={form.ordenCompraSistema}
                onChange={(event) => updateField("ordenCompraSistema", event.target.value)}
                placeholder="OC del sistema"
              />
              {errors.ordenCompraSistema && <small>{errors.ordenCompraSistema}</small>}
            </label>

            <label className="form-field">
              <span>Orden de compra proveedor *</span>
              <input
                type="text"
                value={form.ordenCompraProveedor}
                onChange={(event) => updateField("ordenCompraProveedor", event.target.value)}
                placeholder="OC del proveedor"
              />
              {errors.ordenCompraProveedor && <small>{errors.ordenCompraProveedor}</small>}
            </label>

            <label className="form-field">
              <span>Puerto de salida *</span>
              <select
                value={form.puertoSalida}
                onChange={(event) => updateField("puertoSalida", event.target.value)}
              >
                <option value="">Selecciona una opción</option>
                {puertoSalidaOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.puertoSalida && <small>{errors.puertoSalida}</small>}
            </label>

            <label className="form-field">
              <span>Puerto de llegada *</span>
              <select
                value={form.puertoLlegada}
                onChange={(event) => updateField("puertoLlegada", event.target.value)}
              >
                <option value="">Selecciona una opción</option>
                {puertoLlegadaOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.puertoLlegada && <small>{errors.puertoLlegada}</small>}
            </label>
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>05</span>
            <div>
              <h2>Proveedor</h2>
              <p>Campos P, Q y R del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <div className="form-field form-field-wide">
              <span>Proveedor *</span>

              <div className="proveedor-autocomplete">
                <input
                  type="text"
                  value={form.proveedor}
                  onChange={(event) => {
                    updateField("proveedor", event.target.value);
                    setIsProveedorSuggestionsOpen(true);
                  }}
                  onFocus={() => setIsProveedorSuggestionsOpen(true)}
                  onBlur={() => {
                    window.setTimeout(() => {
                      setIsProveedorSuggestionsOpen(false);
                    }, 150);
                  }}
                  placeholder="Escribe o selecciona un proveedor"
                  autoComplete="off"
                />

                {isProveedorSuggestionsOpen && (
                  <div className="proveedor-suggestions-panel">
                    <div className="proveedor-suggestions-header">
                      <strong>Proveedores registrados</strong>
                      <small>
                        Puedes seleccionar uno o seguir escribiendo manualmente.
                      </small>
                    </div>

                    <div className="proveedor-suggestions-list">
                      {filteredProveedorOptions.map((option) => (
                        <button
                          key={option}
                          type="button"
                          className="proveedor-suggestion-option"
                          onMouseDown={(event) => {
                            event.preventDefault();
                            handleSelectProveedor(option);
                          }}
                        >
                          {option}
                        </button>
                      ))}

                      {filteredProveedorOptions.length === 0 && (
                        <div className="proveedor-suggestion-empty">
                          No se encontraron coincidencias. Puedes capturarlo manualmente.
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {errors.proveedor && <small>{errors.proveedor}</small>}
            </div>

            <label className="form-field form-field-wide">
              <span>Datos fiscales proveedor *</span>
              <textarea
                value={form.datosFiscalesProveedor}
                onChange={(event) => updateField("datosFiscalesProveedor", event.target.value)}
                placeholder="RFC, razón social, domicilio fiscal u otros datos necesarios"
                rows={6}
              />
              {errors.datosFiscalesProveedor && <small>{errors.datosFiscalesProveedor}</small>}
            </label>

            <label className="form-field form-field-wide">
              <span>Información contacto proveedor *</span>
              <textarea
                value={form.informacionContactoProveedor}
                onChange={(event) => updateField("informacionContactoProveedor", event.target.value)}
                placeholder="Nombre, correo, teléfono y observaciones de contacto"
                rows={6}
              />
              {errors.informacionContactoProveedor && (
                <small>{errors.informacionContactoProveedor}</small>
              )}
            </label>
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>06</span>
            <div>
              <h2>Comentarios y registro</h2>
              <p>Campos S y T del archivo Excel.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field form-field-wide">
              <span>Comentarios</span>
              <textarea
                value={form.comentarios}
                onChange={(event) => updateField("comentarios", event.target.value)}
                placeholder="Agrega comentarios adicionales si aplica"
                rows={4}
              />
            </label>

            <label className="form-field form-field-wide">
              <span>Realizado por</span>
              <input
                className="readonly-field"
                type="text"
                value={realizadoPor}
                readOnly
              />
            </label>
          </div>
        </section>

        <section className="form-section">
          <div className="section-header">
            <span>07</span>
            <div>
              <h2>Datos adjuntos</h2>
              <p>Archivos que se enviarán por correo junto con el requerimiento.</p>
            </div>
          </div>

          <div className="form-grid">
            <div className="form-field form-field-wide">
              <span>Archivos adjuntos</span>

              <div className="attachment-email-note">
                <div>
                  <span>Correo destino</span>
                  <strong>{EXPORT_EMAIL_RECIPIENT}</strong>
                </div>
                <small>Al conectar el flujo, los archivos seleccionados se enviarán a este correo.</small>
              </div>

              <div className="attachment-uploader">
                <input
                  id="export-attachments"
                  className="attachment-input"
                  type="file"
                  accept={ACCEPTED_ATTACHMENT_TYPES}
                  multiple
                  onChange={handleAttachmentChange}
                />

                <label className="attachment-dropzone" htmlFor="export-attachments">
                  <strong>Seleccionar archivos</strong>
                  <span>PDF, XLS o XLSX. Máximo 5 archivos, 10 MB por archivo.</span>
                </label>
              </div>

              {attachmentError && <small>{attachmentError}</small>}

              {attachedFiles.length > 0 && (
                <div className="attachment-list">
                  <div className="attachment-list-header">
                    <strong>
                      {attachedFiles.length} de {MAX_ATTACHMENT_FILES} archivos seleccionados
                    </strong>

                    <button type="button" onClick={handleClearAttachments}>
                      Quitar todos
                    </button>
                  </div>

                  {attachedFiles.map((file, index) => (
                    <div key={`${file.name}-${file.size}-${file.lastModified}`} className="attachment-item">
                      <div className="attachment-file-info">
                        <strong>{file.name}</strong>
                        <span>{formatFileSize(file.size)}</span>
                      </div>

                      <button
                        type="button"
                        className="attachment-remove-button"
                        onClick={() => handleRemoveAttachment(index)}
                      >
                        Quitar
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </section>

        {submitMessage && (
          <div className={`submit-message submit-message-${submitMessage.type}`}>
            {submitMessage.text}
          </div>
        )}

        <div className="form-actions">
          <button
            type="button"
            className="secondary-button"
            onClick={handleClear}
            disabled={isSubmitting}
          >
            Limpiar
          </button>

          <button type="submit" className="primary-button" disabled={isSubmitting}>
            {isSubmitting ? "Preparando..." : "Enviar requerimiento"}
          </button>
        </div>
      </form>

      {isConsignatarioModalOpen && (
        <div className="consignatario-modal-overlay">
          <div className="consignatario-modal" role="dialog" aria-modal="true">
            <div className="consignatario-modal-header">
              <h3>CONSIGNATARIO</h3>

              <button
                type="button"
                className="consignatario-modal-close"
                onClick={() => setIsConsignatarioModalOpen(false)}
                aria-label="Cerrar selector de consignatario"
              >
                ×
              </button>
            </div>

            <div className="consignatario-modal-search">
              <input
                type="text"
                value={consignatarioSearch}
                onChange={(event) => setConsignatarioSearch(event.target.value)}
                placeholder="Buscar"
                autoFocus
              />
            </div>

            <div className="consignatario-modal-options">
              {filteredConsignatarioOptions.map((option) => (
                <label key={option} className="consignatario-modal-option">
                  <input
                    type="checkbox"
                    checked={form.consignatario.includes(option)}
                    onChange={() => toggleConsignatario(option)}
                  />
                  <span>{option}</span>
                </label>
              ))}

              {filteredConsignatarioOptions.length === 0 && (
                <div className="consignatario-empty">
                  No se encontraron coincidencias.
                </div>
              )}
            </div>

            <div className="consignatario-modal-footer">
              <button
                type="button"
                className="consignatario-footer-button"
                onClick={handleSelectAllConsignatarios}
              >
                Seleccionar todos
              </button>

              {form.consignatario.length > 0 && (
                <button
                  type="button"
                  className="consignatario-footer-button"
                  onClick={handleClearConsignatarios}
                >
                  Limpiar
                </button>
              )}

              <button
                type="button"
                className="consignatario-done-button"
                onClick={() => setIsConsignatarioModalOpen(false)}
              >
                Hecho
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

export default ExportRequestFormPage;
