import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm";

// Autenticación por NAA (Nested App Authentication): el taskpane pide un
// token real de Entra ID para el usuario que tiene la sesión abierta en
// Outlook, sin ningún secreto compartido embebido aquí (este fichero es
// estático puro en GitHub Pages, no hay servidor propio que lo sirva para
// poder inyectar nada). El backend valida el token en cada llamada — ver
// src/addin/auth.ts en el repo del backend.
const CLIENT_ID = "d6bbe6d2-4287-45fd-b976-86cbdf8047ef";
const TENANT_ID = "3ec777bd-8b86-46a8-800f-6d98eab6bc39";

// Migrado de sage200-mcp a aelis-connect-api (2026-09-29): el backend ya no depende del agente
// de Copilot Studio - código determinista + interpretación con Azure OpenAI, en dos pasos:
// /previsualizar SOLO interpreta y resuelve contra Sage (no crea nada, para que se pueda revisar
// aquí mismo) y /crear escribe el documento en Sage con los datos EXACTOS ya revisados (no
// vuelve a interpretar nada).
const BASE_URL = "https://aelis-connect-api.greenbeach-fdb4a5bf.westeurope.azurecontainerapps.io/addin";
const ADJUNTO_URL = `${BASE_URL}/adjunto`;

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

// Enviamos el ID token, no el access token: el access token de Microsoft
// Graph puede venir cifrado (Microsoft lo documenta como un blob opaco, no
// garantiza que sea un JWT verificable por terceros), mientras que el ID
// token SIEMPRE es un JWT firmado pensado justo para que el propio backend
// identifique al usuario — con la audiencia de NUESTRA app (CLIENT_ID), no
// la de Graph.
//
// OJO: acquireTokenSilent (incluso con forceRefresh) renueva el access
// token pero NO el id_token — MSAL sigue devolviendo el id_token cacheado
// de la autenticación original, que caduca a la hora. Comprobamos su "exp"
// nosotros mismos y forzamos una reautenticación real (popup) si ya venció.
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

let itemActual, nombreRemitenteActual, emailRemitenteActual, adjuntosActuales;
let previsualizacionActual = null; // { cliente, lineas, total } - lo último devuelto por /previsualizar

Office.onReady(() => {
  itemActual = Office.context.mailbox.item;
  const datosDiv = document.getElementById("datos");
  const adjuntosDiv = document.getElementById("adjuntos");
  const botonPrevisualizar = document.getElementById("btnPrevisualizar");
  const botonCrear = document.getElementById("btnCrear");
  const botonVolver = document.getElementById("btnVolver");

  const remitente = itemActual.from || itemActual.sender;
  nombreRemitenteActual = remitente ? remitente.displayName : "";
  emailRemitenteActual = remitente ? remitente.emailAddress : "";

  datosDiv.innerHTML =
    '<div class="kv"><span>De</span><b>' + escapeHtml(nombreRemitenteActual) + "</b></div>" +
    '<div class="kv"><span>Email</span><b>' + escapeHtml(emailRemitenteActual) + "</b></div>" +
    '<div class="kv"><span>Asunto</span><b>' + escapeHtml(itemActual.subject || "") + "</b></div>";

  // isInline descarta las imágenes incrustadas en el cuerpo (logos de firma,
  // etc.): Outlook las expone como adjuntos de tipo File igual que un archivo
  // real, pero no son documentos que el comercial haya adjuntado a mano.
  adjuntosActuales = (itemActual.attachments || []).filter(
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline,
  );
  if (adjuntosActuales.length > 0) {
    adjuntosDiv.innerHTML =
      '<div class="card">' +
      adjuntosActuales.map((a) => '<div class="attachment">📎 ' + escapeHtml(a.name) + "</div>").join("") +
      "</div>";
  }

  botonPrevisualizar.disabled = false;
  botonPrevisualizar.addEventListener("click", previsualizar);
  botonCrear.addEventListener("click", crear);
  botonVolver.addEventListener("click", volverAPrevisualizar);
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

// Reduce el tamaño de las fotos (habitual en correos con tickets/albaranes fotografiados desde
// el móvil) antes de subirlas: sin esto, varias fotos a resolución completa en el mismo correo
// se acercaban al límite de tamaño de petición de la plataforma. No se toca si no es una imagen
// rasterizable (SVG) o si comprimir no reduce el tamaño.
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

function mostrarMensaje(html, clase) {
  document.getElementById("mensaje").innerHTML = `<div class="msg ${clase || ""}">${html}</div>`;
}

function renderPrevisualizacion({ cliente, lineas, total }) {
  const filas = lineas
    .map(
      (l) =>
        `<tr><td>${escapeHtml(l.articulo)}</td><td>${escapeHtml(l.cantidad)}</td><td>${escapeHtml(l.importe.toFixed(2))} €</td></tr>`,
    )
    .join("");
  document.getElementById("previsualizacion").innerHTML =
    `<div class="kv"><span>Cliente</span><b>${escapeHtml(cliente.nombre)}</b></div>` +
    `<div class="kv"><span>Email</span><b>${escapeHtml(cliente.email || "(sin email)")}</b></div>` +
    `<div class="kv"><span>Código Sage</span><b>${escapeHtml(cliente.codigoCliente)}</b></div>` +
    `<table><thead><tr><th>Artículo</th><th>Cant.</th><th>Importe</th></tr></thead><tbody>${filas}</tbody></table>` +
    `<div class="total"><span>Total</span><span>${escapeHtml(total.toFixed(2))} €</span></div>`;
}

async function previsualizar() {
  const boton = document.getElementById("btnPrevisualizar");
  boton.disabled = true;
  boton.textContent = "Leyendo el correo…";
  mostrarMensaje("");

  try {
    const idToken = await acquireIdToken();
    const cuerpoCorreo = await leerCuerpoCorreo(itemActual);

    const adjuntosProcesados = [];
    for (let i = 0; i < adjuntosActuales.length; i++) {
      const adjunto = adjuntosActuales[i];
      mostrarMensaje(`Subiendo adjuntos… (${i + 1}/${adjuntosActuales.length}: ${escapeHtml(adjunto.name)})`);
      try {
        const contenido = await leerAdjunto(itemActual, adjunto.id);
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

    mostrarMensaje("Interpretando el correo…");

    const respuesta = await fetch(`${BASE_URL}/previsualizar`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${idToken}` },
      body: JSON.stringify({
        asunto: itemActual.subject || "",
        cuerpoCorreo: cuerpoCorreo,
        remitente: { nombre: nombreRemitenteActual || emailRemitenteActual, email: emailRemitenteActual },
        adjuntos: adjuntosProcesados,
      }),
    });

    const data = await respuesta.json().catch(() => ({}));
    if (!respuesta.ok) {
      throw new Error(data.error_description || "Error " + respuesta.status);
    }

    previsualizacionActual = data;
    renderPrevisualizacion(data);
    mostrarMensaje("");
    document.getElementById("pasoPrevisualizar").classList.add("oculto");
    document.getElementById("pasoCrear").classList.remove("oculto");
  } catch (err) {
    mostrarMensaje("No se pudo previsualizar: " + escapeHtml(err.message || String(err)), "err");
  } finally {
    boton.disabled = false;
    boton.textContent = "Previsualizar pedido";
  }
}

function volverAPrevisualizar() {
  previsualizacionActual = null;
  mostrarMensaje("");
  document.getElementById("pasoCrear").classList.add("oculto");
  document.getElementById("pasoPrevisualizar").classList.remove("oculto");
}

async function crear() {
  if (!previsualizacionActual) return;
  const boton = document.getElementById("btnCrear");
  const tipo = document.querySelector('input[name="tipoDocumento"]:checked').value;
  boton.disabled = true;
  boton.textContent = "Creando…";
  mostrarMensaje("");

  try {
    const idToken = await acquireIdToken();
    const respuesta = await fetch(`${BASE_URL}/crear`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${idToken}` },
      body: JSON.stringify({ tipo, cliente: previsualizacionActual.cliente, lineas: previsualizacionActual.lineas }),
    });

    const data = await respuesta.json().catch(() => ({}));
    if (!respuesta.ok) {
      throw new Error(data.error_description || "Error " + respuesta.status);
    }

    const textoOk =
      tipo === "presupuesto"
        ? "Presupuesto creado en Sage 200. Se ha enviado un correo de confirmación a " + escapeHtml(previsualizacionActual.cliente.email) + "."
        : "Pedido creado en Sage 200 (documento " + escapeHtml(data.numeroDocumento) + ").";
    mostrarMensaje(textoOk, "ok");
    boton.textContent = "Documento creado";
  } catch (err) {
    mostrarMensaje("No se pudo crear el documento: " + escapeHtml(err.message || String(err)), "err");
    boton.disabled = false;
    boton.textContent = "Crear documento en Sage 200";
  }
}
