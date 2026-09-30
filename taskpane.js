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
// de Copilot Studio - código determinista + interpretación con Azure OpenAI, y responde
// SÍNCRONO (sin jobId/polling) porque interpretar el correo y resolver cliente/artículos contra
// Sage ya es rápido.
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

// El caché de MSAL vive en localStorage del propio origen del taskpane (GitHub Pages) - eso
// SOBREVIVE a cerrar y reabrir Outlook (localStorage no se limpia al cerrar la app), así que si
// el id_token cacheado quedó corrupto/caducado de forma que ni siquiera nuestra propia
// comprobación de "exp" lo detecta bien, reabrir Outlook no arregla nada. Se usa como último
// recurso cuando el backend ya rechazó el token (401) pese a que aquí parecía válido.
function limpiarCacheMsal() {
  Object.keys(localStorage)
    .filter((k) => k.includes(CLIENT_ID) || k.toLowerCase().includes("login.windows.net") || k.startsWith("msal."))
    .forEach((k) => localStorage.removeItem(k));
  msalInstance = undefined;
}

async function acquireIdToken(forzarNuevo) {
  if (forzarNuevo) limpiarCacheMsal();
  await initMsal();
  const tokenRequest = { scopes: ["User.Read"] };
  if (forzarNuevo) {
    const resultado = await msalInstance.acquireTokenPopup(tokenRequest);
    return resultado.idToken;
  }
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

/**
 * POST autenticado con un reintento automático: si el backend responde 401 (id_token caducado
 * o inválido pese a que aquí parecía bueno - ver limpiarCacheMsal), se limpia el caché de MSAL,
 * se fuerza una reautenticación real y se repite la petición UNA vez con el token nuevo.
 * `auth.idToken` se actualiza in situ para que las siguientes llamadas ya usen el token fresco.
 */
async function postConReintentoAuth(url, body, auth) {
  const hacerPeticion = () =>
    fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${auth.idToken}` },
      body: JSON.stringify(body),
    });

  let respuesta = await hacerPeticion();
  if (respuesta.status === 401) {
    auth.idToken = await acquireIdToken(true);
    respuesta = await hacerPeticion();
  }
  const data = await respuesta.json().catch(() => ({}));
  return { respuesta, data };
}

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  const datosDiv = document.getElementById("datos");
  const adjuntosDiv = document.getElementById("adjuntos");
  const mensajeDiv = document.getElementById("mensaje");
  const botonCrear = document.getElementById("btnCrear");
  const botonVisualizar = document.getElementById("btnVisualizar");

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

  const seleccionDiv = document.getElementById("seleccionArticulos");

  botonCrear.disabled = false;
  botonCrear.addEventListener("click", () =>
    crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, botonCrear, botonVisualizar, mensajeDiv, seleccionDiv),
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

async function subirAdjunto(auth, nombreArchivo, contentType, contenidoBase64) {
  const { respuesta, data } = await postConReintentoAuth(ADJUNTO_URL, { nombreArchivo, contentType, contenidoBase64 }, auth);
  if (!respuesta.ok) {
    throw new Error(data.error_description || `No se pudo subir el adjunto ${nombreArchivo} (error ${respuesta.status}).`);
  }
  return data;
}

async function crearPresupuesto(item, nombreRemitente, emailRemitente, adjuntos, boton, botonVisualizar, mensajeDiv, seleccionDiv) {
  boton.disabled = true;
  boton.textContent = "Creando presupuesto…";
  mensajeDiv.innerHTML = "";
  seleccionDiv.innerHTML = "";
  botonVisualizar.classList.add("oculto");

  const auth = { idToken: null };
  try {
    auth.idToken = await acquireIdToken();
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
          const procesado = await subirAdjunto(auth, adjunto.name, contentType, base64);
          adjuntosProcesados.push(procesado);
        }
      } catch (e) {
        console.warn("No se pudo procesar el adjunto", adjunto.name, e);
      }
    }

    mensajeDiv.innerHTML = '<div class="msg">Interpretando el correo…</div>';

    const { respuesta, data } = await postConReintentoAuth(
      `${BASE_URL}/generar-presupuesto`,
      {
        asunto: item.subject || "",
        cuerpoCorreo: cuerpoCorreo,
        remitente: { nombre: nombreRemitente || emailRemitente, email: emailRemitente },
        adjuntos: adjuntosProcesados,
      },
      auth,
    );
    if (!respuesta.ok) {
      throw new Error(data.error_description || "Error " + respuesta.status);
    }

    mensajeDiv.innerHTML = "";
    boton.textContent = "Revisa los datos abajo";
    pintarRevision(auth, data, emailRemitente, boton, botonVisualizar, mensajeDiv, seleccionDiv);
  } catch (err) {
    mensajeDiv.innerHTML =
      '<div class="msg err">No se pudo generar el presupuesto: ' + escapeHtml(err.message || String(err)) + "</div>";
    boton.disabled = false;
    boton.textContent = "Revisar presupuesto";
  }
}

// Pantalla única de revisión antes de escribir nada en Sage: TODOS los artículos (los que se
// resolvieron solos, de solo lectura, y un <select> por cada uno pendiente de elegir a mano) más
// la dirección de entrega interpretada del correo (editable, por si hay que corregirla) -
// directamente en el propio taskpane, sin abrir ninguna pestaña aparte.
function pintarRevision(auth, datos, emailRemitente, boton, botonVisualizar, mensajeDiv, seleccionDiv) {
  const { token, resueltas, pendientes, direccionEntrega } = datos;

  const filasResueltas = resueltas
    .map(
      (l) =>
        `<div class="card linea-pendiente"><div class="original">${escapeHtml(l.descripcion || l.articulo)} — ${escapeHtml(l.cantidad)} unidades</div></div>`,
    )
    .join("");

  const bloquesPendientes = pendientes
    .map((p, i) => {
      const opciones = p.candidatos
        .map(
          (c) =>
            `<option value="${escapeHtml(c.articulo)}">${escapeHtml(c.descripcion)} · ${escapeHtml(c.precioUnitario.toFixed(2))} €</option>`,
        )
        .join("");
      return (
        '<div class="card linea-pendiente">' +
        `<div class="original">${escapeHtml(p.descripcionOriginal)} — ${escapeHtml(p.cantidad)} unidades</div>` +
        `<select data-indice="${i}"><option value="" disabled selected>Elige un artículo…</option>${opciones}</select>` +
        "</div>"
      );
    })
    .join("");

  seleccionDiv.innerHTML =
    filasResueltas +
    bloquesPendientes +
    '<label style="display:block;margin:10px 0 4px;font-size:12.5px;color:var(--muted);">Dirección de entrega' +
    `<textarea id="direccionEntrega" rows="2" style="width:100%;margin-top:4px;padding:8px;border-radius:8px;border:1.5px solid var(--border);font-size:13.5px;font-family:inherit;box-sizing:border-box;">${escapeHtml(direccionEntrega || "")}</textarea>` +
    "</label>" +
    '<button id="btnConfirmarSeleccion" style="margin-top:4px;">Confirmar y crear presupuesto</button>';

  const selects = Array.from(seleccionDiv.querySelectorAll("select"));
  const campoDireccion = document.getElementById("direccionEntrega");
  const botonConfirmar = document.getElementById("btnConfirmarSeleccion");

  botonConfirmar.addEventListener("click", async () => {
    const codigosElegidos = selects.map((s) => s.value);
    if (codigosElegidos.some((c) => !c)) {
      mensajeDiv.innerHTML = '<div class="msg err">Elige un artículo en cada línea antes de continuar.</div>';
      return;
    }

    botonConfirmar.disabled = true;
    botonConfirmar.textContent = "Creando presupuesto…";
    mensajeDiv.innerHTML = "";

    try {
      const { respuesta, data } = await postConReintentoAuth(
        `${BASE_URL}/completar-seleccion`,
        { token, codigosElegidos, direccionEntrega: campoDireccion.value },
        auth,
      );
      if (!respuesta.ok) {
        throw new Error(data.error_description || "Error " + respuesta.status);
      }

      seleccionDiv.innerHTML = "";
      mensajeDiv.innerHTML =
        '<div class="msg ok">Presupuesto creado en Sage 200. Se ha enviado un correo de confirmación a ' +
        escapeHtml(emailRemitente) +
        ".</div>";
      boton.textContent = "Presupuesto creado";

      if (data.urlOferta) {
        botonVisualizar.classList.remove("oculto");
        botonVisualizar.textContent = "Visualizar oferta";
        botonVisualizar.onclick = () => window.open(data.urlOferta, "_blank");
      }
    } catch (err) {
      mensajeDiv.innerHTML =
        '<div class="msg err">No se pudo crear el presupuesto: ' + escapeHtml(err.message || String(err)) + "</div>";
      botonConfirmar.disabled = false;
      botonConfirmar.textContent = "Confirmar y crear presupuesto";
    }
  });
}
