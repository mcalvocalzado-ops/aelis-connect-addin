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

async function acquireAccessToken() {
  await initMsal();
  const tokenRequest = { scopes: ["User.Read"] };
  try {
    const resultado = await msalInstance.acquireTokenSilent(tokenRequest);
    return resultado.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const resultado = await msalInstance.acquireTokenPopup(tokenRequest);
      return resultado.accessToken;
    }
    throw err;
  }
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
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File,
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

async function crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, boton, mensajeDiv) {
  boton.disabled = true;
  boton.textContent = "Creando presupuesto…";
  mensajeDiv.innerHTML = "";

  try {
    const accessToken = await acquireAccessToken();
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
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
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
