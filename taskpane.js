import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm";

// Autenticación por NAA (Nested App Authentication): el taskpane pide un
// token real de Entra ID para el usuario que tiene la sesión abierta en
// Outlook, sin ningún secreto compartido embebido aquí (este fichero es
// estático puro en GitHub Pages, no hay servidor propio que lo sirva para
// poder inyectar nada). El backend valida el token en cada llamada — ver
// src/addin/auth.ts en el repo del backend.
const CLIENT_ID = "d6bbe6d2-4287-45fd-b976-86cbdf8047ef";
const TENANT_ID = "3ec777bd-8b86-46a8-800f-6d98eab6bc39";
const BACKEND_URL = "https://sage200-mcp.greenbeach-fdb4a5bf.westeurope.azurecontainerapps.io/addin/generar-presupuesto";
const ADJUNTO_URL = "https://sage200-mcp.greenbeach-fdb4a5bf.westeurope.azurecontainerapps.io/addin/adjunto";
const ESTADO_URL_BASE = "https://sage200-mcp.greenbeach-fdb4a5bf.westeurope.azurecontainerapps.io/addin/estado/";

let msalInstance;

async function initMsal() {
        if (!msalInstance) {
                  msalInstance = await createNestablePublicClientApplication({
                              auth: {
                                            clientId: CLIENT_ID,
                                            authority: `https://login.microsoftonline.com/${TENANT_ID}`,
                              },
                              cache: { cacheLocation: "localStorage" },
                  });
        }
}

function idTokenCaducado(idTokenClaims) {
        if (!idTokenClaims || typeof idTokenClaims.exp !== "number") return true;
        return idTokenClaims.exp <= Math.floor(Date.now() / 1000);
}

async function acquireIdToken() {
        await initMsal();
        const tokenRequest = { scopes: ["User.Read"] };
        let resultado;
        try {
                  resultado = await msalInstance.acquireTokenSilent(tokenRequest);
                  if (idTokenCaducado(resultado.idTokenClaims)) {
                              resultado = await msalInstance.acquireTokenPopup(tokenRequest);
                  }
        } catch (err) {
                  if (err instanceof InteractionRequiredAuthError) {
                              resultado = await msalInstance.acquireTokenPopup(tokenRequest);
                  } else {
                              throw err;
                  }
        }
        return resultado.idToken;
}

Office.onReady(() => {
        const item = Office.context.mailbox.item;
        const datosDiv = document.getElementById("datos");
        const adjuntosDiv = document.getElementById("adjuntos");
        const mensajeDiv = document.getElementById("mensaje");
        const boton = document.getElementById("btnCrear");

                 const remitente = item.from || item.sender;
        const nombreRemitente = remitente ? remitente.displayName : "";
        const emailRemitente = remitente ? remitente.emailAddress : "";

                 datosDiv.innerHTML =
                           '<div class="kv"><span>De</span><b>' + escapeHtml(nombreRemitente) + "</b></div>" +
                           '<div class="kv"><span>Email</span><b>' + escapeHtml(emailRemitente) + "</b></div>" +
                           '<div class="kv"><span>Asunto</span><b>' + escapeHtml(item.subject || "") + "</b></div>";

                 const adjuntos = (item.attachments || []).filter(
                           (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline,
                         );
        if (adjuntos.length > 0) {
                  adjuntosDiv.innerHTML =
                              '<div class="card">' +
                              adjuntos.map((a) => '<div class="attachment">📎 ' + escapeHtml(a.name) + "</div>").join("") +
                              "</div>";
        }

                 boton.disabled = false;
        boton.addEventListener("click", () =>
                  crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, boton, mensajeDiv),
                                 );
});

function escapeHtml(valor) {
        return String(valor).replace(/[&<>"']/g, function (c) {
                  return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
        });
}

function leerCuerpoCorreo(item) {
        return new Promise((resolve, reject) => {
                  item.body.getAsync(Office.CoercionType.Text, (result) => {
                              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
                              else reject(result.error);
                  });
        });
}

function leerAdjunto(item, attachmentId) {
        return new Promise((resolve, reject) => {
                  item.getAttachmentContentAsync(attachmentId, (result) => {
                              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
                              else reject(result.error);
                  });
        });
}

function esperar(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
}

async function esperarResultado(jobId, mensajeDiv) {
        const intentosMax = 180;
        for (let intento = 1; intento <= intentosMax; intento++) {
                  await esperar(10000);
                  const respuesta = await fetch(ESTADO_URL_BASE + jobId);
                  if (!respuesta.ok) continue;
                  const estado = await respuesta.json();
                  if (estado.estado === "listo") return estado;
                  if (estado.estado === "error") throw new Error(estado.errorMensaje || "No se pudo crear el presupuesto.");
                  mensajeDiv.innerHTML =
                              '<div class="msg">Creando presupuesto en Sage 200… (puede tardar varios minutos si hay documentos que leer)</div>';
        }
        throw new Error("Está tardando más de lo esperado. Puede que se haya creado igualmente — revisa tu correo en unos minutos.");
}

async function subirAdjunto(idToken, nombreArchivo, contentType, contenidoBase64) {
        const respuesta = await fetch(ADJUNTO_URL, {
                  method: "POST",
                  headers: { "Content-Type": "application/json", Authorization: `Bearer ${idToken}` },
                  body: JSON.stringify({ nombreArchivo, contentType, contenidoBase64 }),
        });
        if (!respuesta.ok) {
                  const detalle = await respuesta.json().catch(() => ({}));
                  throw new Error(detalle.error_description || `No se pudo subir el adjunto ${nombreArchivo} (error ${respuesta.status}).`);
        }
        return respuesta.json();
}

function comprimirImagenSiProcede(base64Original, contentType) {
        return new Promise((resolve) => {
                  if (!contentType.startsWith("image/") || contentType === "image/svg+xml") {
                              resolve({ base64: base64Original, contentType });
                              return;
                  }
                  const MAX_LADO = 1600;
                  const CALIDAD = 0.75;
                  const img = new Image();
                  img.onload = () => {
                              let { width, height } = img;
                              if (width > MAX_LADO || height > MAX_LADO) {
                                            const ratio = Math.min(MAX_LADO / width, MAX_LADO / height);
                                            width = Math.round(width * ratio);
                                            height = Math.round(height * ratio);
                              }
                              const canvas = document.createElement("canvas");
                              canvas.width = width;
                              canvas.height = height;
                              const ctx = canvas.getContext("2d");
                              ctx.drawImage(img, 0, 0, width, height);
                              const dataUrl = canvas.toDataURL("image/jpeg", CALIDAD);
                              const base64Comprimido = dataUrl.split(",")[1] || "";
                              if (base64Comprimido && base64Comprimido.length < base64Original.length) {
                                            resolve({ base64: base64Comprimido, contentType: "image/jpeg" });
                              } else {
                                            resolve({ base64: base64Original, contentType });
                              }
                  };
                  img.onerror = () => resolve({ base64: base64Original, contentType });
                  img.src = `data:${contentType};base64,${base64Original}`;
        });
}

async function crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, boton, mensajeDiv) {
        boton.disabled = true;
        boton.textContent = "Creando presupuesto…";
        mensajeDiv.innerHTML = '<div class="msg">Leyendo el correo…</div>';

  try {
            const idToken = await acquireIdToken();
            const cuerpoCorreo = await leerCuerpoCorreo(item);

          const adjuntosProcesados = [];
            for (let i = 0; i < adjuntos.length; i++) {
                        const adjunto = adjuntos[i];
                        mensajeDiv.innerHTML =
                                      `<div class="msg">Subiendo adjuntos… (${i + 1}/${adjuntos.length}: ${escapeHtml(adjunto.name)})</div>`;
                        try {
                                      const contenido = await leerAdjunto(item, adjunto.id);
                                      if (contenido.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                                                      const contentTypeOriginal = adjunto.contentType || "application/octet-stream";
                                                      const { base64, contentType } = await comprimirImagenSiProcede(contenido.content, contentTypeOriginal);
                                                      const procesado = await subirAdjunto(idToken, adjunto.name, contentType, base64);
                                                      adjuntosProcesados.push(procesado);
                                      }
                        } catch (e) {
                                      console.warn("No se pudo procesar el adjunto", adjunto.name, e);
                        }
            }

          mensajeDiv.innerHTML = '<div class="msg">Enviando datos a Sage 200…</div>';

          const respuesta = await fetch(BACKEND_URL, {
                      method: "POST",
                      headers: { "Content-Type": "application/json", Authorization: `Bearer ${idToken}` },
                      body: JSON.stringify({
                                    asunto: item.subject || "",
                                    cuerpoCorreo: cuerpoCorreo,
                                    remitente: { nombre: nombreRemitente || emailRemitente, email: emailRemitente },
                                    adjuntos: adjuntosProcesados,
                      }),
          });

          if (!respuesta.ok) {
                      const detalle = await respuesta.json().catch(() => ({}));
                      throw new Error(detalle.error_description || "Error " + respuesta.status);
          }

          const { jobId } = await respuesta.json();
            mensajeDiv.innerHTML = '<div class="msg">Creando presupuesto en Sage 200…</div>';
            await esperarResultado(jobId, mensajeDiv);

          mensajeDiv.innerHTML =
                      '<div class="msg ok">Presupuesto creado en Sage 200. Se ha enviado un correo de confirmación a ' +
                      escapeHtml(emailRemitente) +
                      ".</div>";
            boton.textContent = "Presupuesto creado";
  } catch (err) {
            mensajeDiv.innerHTML =
                        '<div class="msg err">No se pudo crear el presupuesto: ' + escapeHtml(err.message || String(err)) + "</div>";
            boton.disabled = false;
            boton.textContent = "Crear presupuesto en Sage 200";
  }
}
