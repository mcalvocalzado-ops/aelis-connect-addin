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

  // isInline descarta las imágenes incrustadas en el cuerpo (logos de firma,
  // etc.): Outlook las expone como adjuntos de tipo File igual que un archivo
  // real, pero no son documentos que el comercial haya adjuntado a mano.
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

// Poll del estado del trabajo en segundo plano: leer PDFs/imágenes reales
// con el agente puede tardar bastante más que el límite de respuesta
// síncrona de Power Automate, así que el backend responde al momento con un
// jobId y aquí preguntamos periódicamente hasta que termine.
async function esperarResultado(jobId, mensajeDiv) {
  const intentosMax = 40; // 40 x 5s = hasta ~3'20" de margen
  for (let intento = 1; intento <= intentosMax; intento++) {
    await esperar(5000);
    const respuesta = await fetch(ESTADO_URL_BASE + jobId);
    if (!respuesta.ok) continue;
    const estado = await respuesta.json();
    if (estado.estado === "listo") return estado;
    if (estado.estado === "error") throw new Error(estado.errorMensaje || "No se pudo crear el presupuesto.");
    mensajeDiv.innerHTML =
      '<div class="msg">Creando presupuesto en Sage 200… (puede tardar 1-3 minutos si hay documentos que leer)</div>';
  }
  throw new Error("Está tardando más de lo esperado. Puede que se haya creado igualmente — revisa tu correo en unos minutos.");
}

async function crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, boton, mensajeDiv) {
  boton.disabled = true;
  boton.textContent = "Creando presupuesto…";
  mensajeDiv.innerHTML = '<div class="msg">Leyendo el correo…</div>';

  try {
    const idToken = await acquireIdToken();
    const cuerpoCorreo = await leerCuerpoCorreo(item);

    const adjuntosLeidos = [];
    for (const adjunto of adjuntos) {
      try {
        const contenido = await leerAdjunto(item, adjunto.id);
        if (contenido.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          adjuntosLeidos.push({
            nombreArchivo: adjunto.name,
            contentType: adjunto.contentType || "application/octet-stream",
            contenidoBase64: contenido.content,
          });
        }
      } catch (e) {
        console.warn("No se pudo leer el adjunto", adjunto.name, e);
      }
    }

    const respuesta = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${idToken}` },
      body: JSON.stringify({
        asunto: item.subject || "",
        cuerpoCorreo: cuerpoCorreo,
        remitente: { nombre: nombreRemitente || emailRemitente, email: emailRemitente },
        adjuntos: adjuntosLeidos,
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
