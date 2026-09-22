// ==UserScript==
// @name         Qintek - Autocompletar Orden de Pago (Duplicar)
// @namespace    qintek-pmo-automation
// @version      2.8
// @description  Completa SOLO los campos que falten en Captura > Órdenes de pago de Qintek, después de que subas el XML de la factura. Busca la orden a clonar por número de folio dentro del Excel que pegues. Nunca presiona "Guardar".
// @match        https://qintek.qin.mx/crud/capturar/ordenesdepago*
// @grant        GM_registerMenuCommand
// @updateURL    https://raw.githubusercontent.com/jmoreno-alt/qintek-tampermonkey-scripts/main/qintek-autocompletar-ordenes-pago.user.js
// @downloadURL  https://raw.githubusercontent.com/jmoreno-alt/qintek-tampermonkey-scripts/main/qintek-autocompletar-ordenes-pago.user.js
// ==/UserScript==

/*
  FLUJO COMPLETO (acordado con el usuario, v2.0)
  1. Usuario: inicia sesión en Qintek manualmente.
  2. Usuario: entra a Qintek > Captura > Órdenes de pago.
  3. Qintek: el usuario sube el archivo XML (CFDI) de la factura en el propio formulario
     de Qintek. Qintek lo lee y llena automáticamente: Proveedor, Cuenta Destino,
     Folio de Factura, Fecha factura, Confirmar Monto Factura (incluye IVA), IVA Factura,
     Tipo Moneda y Folio UUID.
  4. Este script: detonado con el "número de orden de pago" (folio) y los datos del Excel
     de la orden que se está clonando, llena SOLO lo que quedó faltante después del XML:
     típicamente Empresa Pagadora, Concepto(s) y montos, Método/Forma de Pago,
     Semana Gasto, Centro de Costos (con montos) y Comentarios. Si algún campo que
     debería haber llenado el XML sigue vacío (por ejemplo Qintek no encontró el
     Proveedor), el script lo llena como respaldo usando el Excel.
  5. Usuario: revisa y edita a mano lo que haga falta.
  6. Usuario: da clic en "Guardar". Este script JAMÁS hace ese clic por ti, bajo
     ninguna circunstancia, en ninguna versión presente o futura.

  CÓMO INSTALARLO
  1. Instala la extensión "Tampermonkey" en Chrome (chrome://extensions o la Chrome Web Store).
  2. Abre el panel de Tampermonkey > "Crear nuevo script" (o edita este mismo script si ya
     lo tenías instalado de una versión anterior).
  3. Borra el contenido y pega TODO este archivo.
  4. Guarda (Ctrl+S / ícono de guardar).
  5. Entra a https://qintek.qin.mx/crud/capturar/ordenesdepago — debe aparecer un panel
     naranja flotante "Autocompletar Orden de Pago" en la esquina superior derecha.

  CÓMO USARLO
  1. Sube primero el XML de la factura en el formulario de Qintek (paso nativo de Qintek,
     no de este script) y deja que Qintek llene lo que le corresponde.
  2. En Excel, copia la fila de encabezados Y TODAS las filas de órdenes que podrías
     necesitar clonar (puedes pegar tu Excel completo de solicitudes, no solo una fila).
     Pégalas (Ctrl+V) en el cuadro de texto grande del panel.
  3. Escribe el "Número de orden de pago" (folio) de la orden que estás clonando en el
     campo correspondiente del panel, para que el script sepa cuál fila usar. Si solo
     pegaste una fila de datos, este campo es opcional.
  4. Si el Excel NO trae el nombre del Proveedor (solo la CLABE) y Qintek tampoco lo pudo
     resolver desde el XML, escríbelo tal cual aparece en el catálogo de Qintek en el
     campo "Proveedor" del panel.
  5. Da clic en "Buscar folio y completar faltantes".
  6. Revisa el resumen: ✅ = lo llenó el script, ⏭️ = ese campo ya lo había llenado Qintek
     con el XML (no se tocó, para no pisar lo que trajo la factura), ⚠️/❌ = necesita que
     lo completes o corrijas tú a mano.
  7. Revisa el formulario completo y da clic en "Guardar" TÚ MISMO.

  LÍMITES CONOCIDOS (honestos, para que no los des por sentado)
  - Si Qintek no logró resolver el Proveedor desde el XML y el Excel solo tiene la CLABE
    (sin nombre), el script tampoco puede adivinarlo: el catálogo de Qintek no muestra la
    CLABE en pantalla, solo el nombre. Escríbelo a mano en el campo "Proveedor" del panel.
  - Si un mismo nombre de Centro de Costos existe repetido en más de un tipo, la regla
    (acordada, v2.7) es: si "Planta" es una de las opciones, se usa Planta
    automáticamente (ej. "Om Ah Hum" existe en Sucursales y Planta → se usa Planta).
    Si la ambigüedad NO incluye Planta (por ejemplo "Mantenimiento" como Departamento Y
    como Sucursal), el script no adivina cuál es: lo deja marcado como ⚠️ para que lo
    agregues tú manualmente.
  - El formato de fecha esperado (para el respaldo de "Fecha factura", si el XML no la
    trajo) es día/mes/año, como lo exporta Qintek (ej. "15/9/2026, 12:21:34 p.m.").
  - Si una misma orden trae Centro de Costos de los 3 tipos A LA VEZ (Departamentos +
    Sucursales + Planta simultáneamente), los del primer tipo que se procesó se pueden
    perder al llegar al último tipo (Qintek borra la selección de "Centro costos" del
    tipo que se va desmarcando, y marcar los 3 tipos a la vez dispara el comportamiento
    especial "-TODOS-" de Qintek, así que no se pueden dejar los 3 marcados al mismo
    tiempo). Si tu orden mezcla los 3 tipos, revisa el resultado final antes de guardar
    y agrega a mano lo que haga falta. Cuando son 1 o 2 tipos distintos (el caso más
    común, incluso con muchas líneas alternando entre ambos), el script marca de una
    vez los tipos que va a necesitar y los deja marcados durante todo el proceso, así
    que no hay pérdida de datos (corregido en v2.5).
  - El script decide "¿ya está lleno?" leyendo lo que se ve en pantalla en cada campo en
    el momento de dar clic en el botón. Si subes el XML DESPUÉS de correr el script, vas
    a ver que rellenó campos que el XML iba a llenar de todos modos: en ese caso, vuelve a
    correr el script (o edita a mano) — no pasa nada, nunca sobrescribe lo que ya tiene
    valor, solo completa lo que está vacío.

  CAMBIOS v1.1
  - Corregido: en "Centro de Costos", la versión anterior usaba el filtro "-TODOS-" en
    "Tipo Centro costos" para ver Departamentos y Sucursales combinados. Se detectó que
    en Qintek esa opción NO solo filtra la lista: hace que la aplicación seleccione
    automáticamente TODOS los centros de costo del catálogo, dejando decenas marcados
    sin monto. Ahora el script revisa Departamentos, Sucursales y Planta por separado
    (sin tocar nunca "-TODOS-") y solo marca el centro de costo que de verdad corresponde
    a cada línea del Excel.

  CAMBIOS v1.2
  - Corregido (de nuevo) "Centro de Costos": "Tipo Centro costos" resultó ser también un
    multiselect, no un dropdown de una sola opción. Al revisar Departamentos, Sucursales
    y Planta uno tras otro, cada clic SUMABA esa opción a las anteriores en vez de
    reemplazarla, así que terminaban los 3 tipos marcados a la vez (y Qintek, al verlos
    los 3 marcados, prendía también "-TODOS-" solo) — mismo síntoma que el bug de v1.0
    pero por una causa distinta. Ahora se desmarca cada tipo justo después de leer su
    catálogo, así que nunca hay más de un tipo marcado al mismo tiempo.
  - Más robustez general: abrir un dropdown/multiselect ahora reintenta hasta 4 veces
    con más espera cada vez (antes solo lo intentaba una vez con 350ms). Si tu conexión
    o la respuesta de Qintek es más lenta, esto debería reducir los "no se encontró la
    opción" en Empresa Pagadora, Proveedor, Concepto, Método/Forma de Pago, etc.

  CAMBIOS v1.3 (causa raíz real, verificada en vivo)
  - Se encontró el motivo de fondo detrás de v1.1 y v1.2: en esta página de Qintek, los
    paneles emergentes de los dropdown/multiselect se agregan todos a <body> y casi nunca
    se eliminan del DOM al cerrarlos — solo quedan ocultos o se acumulan varios a la vez.
    Por eso "tomar el primer panel que exista en la página" (lo que hacían las versiones
    anteriores) a veces devolvía el panel de OTRO campo por error — por ejemplo, pedir el
    panel de "Centro costos" y recibir el de "Tipo Centro costos" que había quedado pegado
    — y de ahí el síntoma de "se marcan centros de costo que no debían".
  - Ahora, antes de abrir cualquier panel se eliminan del DOM los paneles viejos, y si aun
    así hay varios abiertos a la vez se elige el que está alineado con el control en el
    que se acaba de hacer clic (no el primero que aparezca en el HTML). Esto se probó en
    vivo contra Qintek y quedó resuelto.

  CAMBIOS v2.0 (replanteamiento del flujo, a partir de la nueva carga de XML en Qintek)
  - El usuario detectó que Qintek puede leer el XML de la factura directamente en el
    formulario y llenar solo con eso: Proveedor, Cuenta Destino, Folio de Factura, Fecha
    factura, Confirmar Monto Factura, IVA Factura, Tipo Moneda y Folio UUID. El script ya
    no debe volver a llenar esos campos si Qintek ya lo hizo — solo debe completar lo que
    quede faltante.
  - Nuevo: cada uno de esos campos se revisa antes de tocarlo (¿ya tiene un valor visible
    en pantalla?). Si ya tiene valor, el script lo deja intacto y lo marca como ⏭️ en el
    resumen. Si sigue vacío (por ejemplo Qintek no encontró el Proveedor), lo llena como
    respaldo con los datos del Excel, igual que antes.
  - Nuevo: el cuadro de texto ahora acepta pegar VARIAS filas a la vez (encabezados +
    todas las órdenes que podrías clonar, por ejemplo tu Excel completo de solicitudes),
    en vez de solo encabezado+1 fila. Se agregó el campo "Número de orden de pago" para
    que el script busque la fila correcta por folio en vez de asumir que solo hay una.
  - El botón cambió de "Autocompletar formulario" a "Buscar folio y completar faltantes"
    para reflejar el nuevo rol del script (completar lo faltante, no llenar todo desde
    cero).

  CAMBIOS v2.1
  - El panel ahora se llama "Panel de Clonado de Orden de Pago" (nombre acordado con el
    usuario para usarlo también en la documentación del proyecto).
  - Corregido: en una app Angular como Qintek, navegar entre pantallas del menú lateral
    NO siempre recarga la página (routing del lado del cliente), y Tampermonkey solo
    inyecta el script cuando la página se carga de verdad. Si el panel dejaba de verse
    tras navegar dentro de Qintek, antes la única solución era refrescar (F5). Ahora hay
    dos soluciones:
      1) El script vuelve a intentar crear el panel cada segundo mientras la pantalla de
         Órdenes de pago esté abierta (antes solo lo intentaba una vez al cargar).
      2) Se agregó un comando en el menú de Tampermonkey ("Mostrar Panel de Clonado"):
         da clic en el ícono de Tampermonkey en la barra de Chrome mientras estás en la
         pantalla de Qintek y selecciona esa opción para volver a mostrar el panel en
         cualquier momento, sin necesidad de refrescar la página.

  CAMBIOS v2.2 (bug real: varios Centros de Costos en la misma corrida)
  - Con un solo Centro de Costos el script funcionaba bien, pero al pegar varias líneas
    de Centro de Costos (ej. "4-Chapalita: $5162, 37-Centro: $6960, 45-Iteso: $5162")
    dejaba de procesarlos. Causa: al SELECCIONAR de verdad cada Centro de Costos (no al
    armar el catálogo, que ya se corrigió en v1.2), el script marcaba "Tipo Centro
    costos" pero nunca lo volvía a desmarcar después. Por ser un multiselect, esto
    provocaba que: si la siguiente línea era del mismo tipo, el script terminara
    DESMARCÁNDOLO sin querer (un clic sobre una opción ya marcada la quita); y si era de
    un tipo distinto, quedaran dos tipos marcados a la vez, mezclando el catálogo de
    Centro de Costos. Ahora, igual que al armar el catálogo, se desmarca "Tipo Centro
    costos" inmediatamente después de seleccionar cada línea, así que siempre queda como
    máximo un tipo marcado entre una línea y la siguiente.

  CAMBIOS v2.3 (la corrección de v2.2 resultó incompleta — probado en vivo)
  - Al probar v2.2 con varios Centros de Costos del mismo tipo (ej. 3 Sucursales), el
    log mostraba "✅ Centro de Costos agregado" para los tres, pero el formulario
    terminaba con "Tipo Centro costos" y "Centro costos" completamente vacíos y Monto
    Total en $0.00. Causa real: desmarcar "Tipo Centro costos" hasta dejarlo en CERO
    tipos marcados (lo que hacía v2.2 después de cada línea) hace que Qintek borre TODO
    lo que hubiera en "Centro costos" — no solo la línea recién agregada. v2.2
    solucionó el bug de "se desmarca sin querer entre líneas del mismo tipo" pero
    introdujo uno peor al limpiar el campo por completo al final.
  - Ahora: (a) "Tipo Centro costos" solo se toca cuando la siguiente línea de verdad
    necesita un tipo distinto al actual (nunca entre líneas del mismo tipo, que es el
    caso más común), (b) al cambiar de tipo se marca el nuevo ANTES de desmarcar el
    anterior, para no pasar nunca por "cero tipos marcados", y (c) al terminar se deja
    marcado el último tipo usado en vez de limpiarlo. Limitación conocida que queda:
    si una misma orden mezcla Centro de Costos de los 3 tipos a la vez, los del primer
    tipo procesado se pueden perder (ver LÍMITES CONOCIDOS) — con 1 o 2 tipos distintos
    (el caso normal) queda correcto.

  CAMBIOS v2.4 (bug real, visto en vivo — nombres cortos de Centro de Costos)
  - v2.3 ya seleccionaba bien varios Centro de Costos del mismo tipo, pero probando con
    "Chapalita", "Centro" e "Iteso" el monto de "Centro" quedó vacío (Monto Total
    incompleto: $10,324 en vez de $17,284). Causa: la búsqueda del campo de monto
    (`findRowAmountInput`) hacía coincidir por "el texto de la fila EMPIEZA CON el
    nombre". "Centro" es tan corto que también "empieza" el título de la sección
    ("Centro de Costos") y la etiqueta del campo ("Centro costos"), que además
    aparecen ANTES en el HTML que la fila real "Centro (Sucursales)" — así que el
    código se quedaba con esa coincidencia equivocada en vez de con la fila correcta.
    Con "Chapalita" e "Iteso" no pasaba por ser nombres únicos en toda la pantalla.
  - Ahora se guarda el texto EXACTO de cada Centro de Costos ya seleccionado (ej.
    "Centro (Sucursales)", con su sufijo de tipo) y se busca su input de monto por
    coincidencia EXACTA primero; "empieza con" solo se usa como respaldo (por ejemplo
    para las filas de Concepto, que no llevan sufijo de tipo).

  CAMBIOS v2.5 (bug real, visto en vivo — 24 Centro de Costos mezclando 2 tipos)
  - Probando con una orden real de 24 líneas de Centro de Costos que alternaban todo
    el tiempo entre Departamentos y Sucursales (no en bloques, sino mezcladas en el
    orden del Excel), el formulario terminó con UN SOLO Centro de Costos marcado (el
    último de la lista) y el resto perdido — a pesar de que v2.3/v2.4 ya evitaban
    "cero tipos marcados". Causa real: la limitación que hasta ahora se creía exclusiva
    de mezclar los 3 tipos a la vez también aplica cada vez que el script CAMBIA de un
    tipo a otro en cualquier momento del proceso (no solo al final): al desmarcar el
    tipo anterior para marcar el siguiente, Qintek borra TODOS los Centro de Costos ya
    agregados que pertenecían a ese tipo que se está desmarcando — no solo cuando se
    llega a cero tipos. Como esta orden iba cambiando de tipo en casi cada línea, cada
    cambio borraba lo ya avanzado, y solo sobrevivió lo agregado después del último
    cambio.
  - Corregido de raíz: ahora el script primero resuelve a qué tipo pertenece CADA línea
    (usando el mismo catálogo de siempre) y arma el conjunto de tipos realmente
    necesarios ANTES de tocar nada. Si son 1 o 2 tipos (lo normal, sin importar cuántas
    líneas o en qué orden vengan mezcladas), los marca TODOS de una sola vez al
    principio y ya no los vuelve a tocar durante el resto del proceso — así nunca se
    dispara el borrado por cambio de tipo. Solo si de verdad se necesitan los 3 tipos
    a la vez (Departamentos + Sucursales + Planta) se usa el método anterior
    (secuencial, con la pérdida conocida del primer tipo procesado), porque marcar los
    3 al mismo tiempo activa el comportamiento especial "-TODOS-" de Qintek y no se
    puede evitar de otra forma. En ese caso el script ahora avisa explícitamente en el
    resumen antes de empezar.

  CAMBIOS v2.6 (dos bugs reales, vistos en vivo — orden de 24 líneas mezclando
  Departamentos, Sucursales y Planta)
  - Bug 1: TODAS las líneas de "Sucursales" (14 de 24) se reportaron como "no
    encontré ... en Departamentos, Sucursales ni Planta", aunque sí existían en
    Qintek. Causa: al construir el catálogo de Centro de Costos por tipo, el script
    abría el panel y leía sus opciones casi de inmediato — pero las opciones tardan
    en llegar, sobre todo en catálogos largos como "Sucursales" (con más ramas que
    "Departamentos" o "Planta"). El catálogo de Sucursales se leyó vacío o
    incompleto, así que ninguna línea de ese tipo pudo resolverse. Corregido: ahora
    se espera a que el número de opciones del panel se "estabilice" (deje de crecer)
    antes de leerlas, tanto al construir el catálogo como al reabrir el panel para
    seleccionar cada línea.
  - Bug 2: de los Centro de Costos que SÍ se lograron agregar, a los primeros de la
    lista no se les pudo llenar el monto ("no encontré el campo de monto"), mientras
    que a los últimos agregados sí. Causa probable: la lista de filas en pantalla no
    mantiene renderizadas todas las filas ya agregadas cuando hay muchas — por eso
    fallaban las más antiguas. Corregido: ya no se agregan primero TODAS las líneas y
    se rellenan los montos hasta el final; ahora cada línea se agrega Y se le llena
    el monto de inmediato, una por una, mientras su fila recién creada sigue
    garantizada en el DOM.
  - Además: se subieron de 4 a 6 los reintentos para abrir cualquier panel de
    dropdown/multiselect (con más espera entre cada uno), porque también se vio en
    vivo que "Empresa Pagadora" —el primer campo que toca el script en cada
    corrida— a veces no alcanzaba a abrir su panel a tiempo.

  CAMBIOS v2.7 (regla acordada con el usuario)
  - Al corregir el bug del catálogo de Sucursales en v2.6, salió a la luz un caso real
    de ambigüedad: "Om Ah Hum" existe tanto en Sucursales como en Planta, así que el
    script lo marcaba como ⚠️ para agregarlo manualmente. El usuario pidió que, en ese
    tipo de ambigüedad, se prefiera "Planta". Ahora: si un nombre existe en más de un
    tipo y "Planta" es una de las opciones, el script usa Planta automáticamente (y lo
    deja anotado en el resumen). Si la ambigüedad no incluye Planta, se mantiene el
    comportamiento anterior (⚠️, agregar a mano), porque no hay una regla definida para
    esos casos.

  CAMBIOS v2.8 (bug real, visto en vivo — la regla de v2.7 podía arruinar TODO el lote)
  - Al probar v2.7 con la orden real de 24 líneas (que solo necesitaba Departamentos +
    Sucursales — el caso seguro de 2 tipos), la ambigüedad de "Om Ah Hum" se resolvió a
    Planta como se pidió, pero eso hizo que el lote pasara a necesitar 3 tipos a la vez
    — y ahí sí aplica la limitación grave de Qintek (cada cambio de tipo borra lo ya
    agregado del tipo que se desmarca), poniendo en riesgo TODAS las 24 líneas, no solo
    la ambigua. El resumen mostraba ✅ en todas (porque se registra el clic al
    momento), pero el formulario real probablemente terminó con solo una fracción de
    las líneas.
  - Corregido: la preferencia por "Planta" ahora solo se aplica cuando NO implica abrir
    un tercer tipo en el lote (o sea, cuando Planta ya era necesaria de todos modos, o
    el lote de por sí solo necesitaba 1 tipo más). Si forzar Planta abriría un tercer
    tipo, el script en su lugar usa el otro tipo de la ambigüedad si ya era necesario
    (evitando el riesgo por completo), o —si ni eso alcanza— lo deja para agregarse a
    mano, avisando explícitamente que se evitó a propósito para no arriesgar el resto
    del lote.
*/

(function () {
  'use strict';

  // ---------------------------------------------------------------------
  // Utilidades de bajo nivel para manipular el formulario Angular/PrimeNG
  // ---------------------------------------------------------------------

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function setAngularInputValue(input, value) {
    const proto =
      input.tagName === 'TEXTAREA'
        ? window.HTMLTextAreaElement.prototype
        : window.HTMLInputElement.prototype;
    const nativeSetter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    nativeSetter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new Event('blur', { bubbles: true }));
  }

  function findFieldContainer(labelText) {
    const labels = Array.from(document.querySelectorAll('label')).filter(
      (l) => l.textContent.trim() === labelText
    );
    if (!labels.length) return null;
    return labels[0].closest('.field') || labels[0].parentElement;
  }

  function getFieldInput(labelText) {
    const field = findFieldContainer(labelText);
    return field ? field.querySelector('input') : null;
  }

  // NUEVO v2.0: para decidir si un campo "ya lo llenó Qintek con el XML" y por lo
  // tanto debe dejarse intacto, o si sigue vacío y el script debe completarlo.

  function isPlainInputEmpty(labelText) {
    const input = getFieldInput(labelText);
    return !input || !input.value || !input.value.trim();
  }

  function isDropdownEmpty(labelText) {
    const field = findFieldContainer(labelText);
    if (!field) return true;
    const labelEl = field.querySelector('.p-dropdown-label, .p-multiselect-label');
    if (!labelEl) return true;
    if (labelEl.classList.contains('p-placeholder')) return true;
    const text = labelEl.textContent.trim();
    return !text || /seleccion/i.test(text);
  }

  // DESCUBRIMIENTO IMPORTANTE (v1.3): en esta página de Qintek, los paneles de
  // los dropdown/multiselect (PrimeNG) se agregan todos a <body> y MUCHAS VECES
  // NO se eliminan del DOM al cerrarlos con un clic afuera — solo quedan
  // visualmente ocultos o directamente se acumulan varios a la vez. Por eso
  // "tomar el primer .p-multiselect-panel que exista" (lo que hacían v1.0–v1.2)
  // podía devolver el panel de OTRO campo (por ejemplo, pedir el panel de
  // "Centro costos" y recibir por error el de "Tipo Centro costos" que había
  // quedado pegado), y de ahí varios de los síntomas raros reportados. La
  // solución: 1) purgar del DOM cualquier panel viejo antes de abrir uno nuevo,
  // y 2) si aun así hay varios paneles visibles, quedarnos con el que está
  // alineado horizontalmente con el control en el que acabamos de hacer clic.

  function getAllPanels() {
    return Array.from(
      document.querySelectorAll('.p-multiselect-panel, .p-dropdown-panel')
    ).filter((p) => p.offsetParent !== null);
  }

  function purgeStalePanels() {
    document
      .querySelectorAll('.p-multiselect-panel, .p-dropdown-panel')
      .forEach((p) => p.remove());
  }

  function findPanelNearTrigger(triggerEl) {
    const panels = getAllPanels();
    if (!panels.length) return null;
    const tRect = triggerEl.getBoundingClientRect();
    let best = null;
    let bestOverlap = -1;
    for (const p of panels) {
      const r = p.getBoundingClientRect();
      const overlap = Math.min(r.right, tRect.right) - Math.max(r.left, tRect.left);
      if (overlap > bestOverlap) {
        bestOverlap = overlap;
        best = p;
      }
    }
    return best;
  }

  function getTrigger(labelText) {
    const field = findFieldContainer(labelText);
    if (!field) throw new Error(`No se encontró el campo "${labelText}"`);
    const trigger = field.querySelector('.p-dropdown, .p-multiselect');
    if (!trigger) throw new Error(`No se encontró el control de "${labelText}"`);
    return trigger;
  }

  // Reintenta abrir el panel: en Qintek a veces el primer clic no alcanza a
  // abrir el overlay (animación/lag), así que probamos varias veces con más
  // espera cada vez antes de rendirnos. Siempre purga paneles viejos primero
  // y elige el panel alineado con el control correcto (ver nota arriba).
  // v2.6: se subió de 4 a 6 intentos y se alargó la espera de cada uno — se vio en vivo
  // que "Empresa Pagadora" (el primer campo que toca el script en cada corrida) a veces
  // no alcanza a abrir su panel con el límite anterior, probablemente porque Qintek
  // todavía está terminando de asentarse justo después de procesar el XML.
  async function openFieldPanel(labelText, attempts = 6) {
    const trigger = getTrigger(labelText);
    purgeStalePanels();
    for (let i = 0; i < attempts; i++) {
      trigger.click();
      await sleep(450 + i * 300);
      const panel = findPanelNearTrigger(trigger);
      if (panel) return panel;
    }
    return null;
  }

  // v2.6 (bug real, visto en vivo — catálogo de "Sucursales" vacío): al construir el
  // catálogo de Centro de Costos por tipo, el panel aparecía en el DOM casi de
  // inmediato pero sus opciones (<li>) seguían llegando unos cientos de milisegundos
  // después — sobre todo para catálogos largos como "Sucursales", que tiene muchas más
  // opciones que "Departamentos" o "Planta". El código anterior leía las opciones justo
  // al abrir el panel y a veces se quedaba con una lista vacía o incompleta para ese
  // tipo, lo que hacía fallar TODAS las líneas de Centro de Costos de ese tipo (se
  // reportaban como "no encontrado en Departamentos, Sucursales ni Planta" aunque sí
  // existieran). Ahora se espera a que el número de opciones se "estabilice" (deje de
  // crecer entre una revisión y la siguiente) antes de leerlas.
  async function waitForStableOptions(panel, stepMs = 200, maxAttempts = 8) {
    let prevCount = -1;
    for (let i = 0; i < maxAttempts; i++) {
      const count = panel.querySelectorAll('li').length;
      if (count > 0 && count === prevCount) return;
      prevCount = count;
      await sleep(stepMs);
    }
  }

  // Selecciona UNA opción por texto (sirve tanto para dropdown simple como
  // para multiselect: si el panel sigue abierto tras seleccionar, lo cierra).
  async function selectDropdownOption(labelText, optionText) {
    const trigger = getTrigger(labelText);
    let panel = await openFieldPanel(labelText);
    if (!panel) throw new Error(`No abrió el panel de "${labelText}"`);

    const filterInput = panel.querySelector(
      '.p-dropdown-filter, .p-multiselect-filter-container input'
    );
    if (filterInput) {
      setAngularInputValue(filterInput, optionText);
      await sleep(350);
      panel = findPanelNearTrigger(trigger);
      if (!panel) throw new Error(`El panel de "${labelText}" se cerró al filtrar`);
    }

    const items = Array.from(panel.querySelectorAll('li'));
    const norm = (s) => s.trim().toLowerCase();
    const exact = items.find((li) => norm(li.textContent) === norm(optionText));
    const partial = items.find((li) => norm(li.textContent).includes(norm(optionText)));
    const match = exact || partial;

    if (!match) {
      document.body.click();
      throw new Error(`No se encontró la opción "${optionText}" en "${labelText}"`);
    }
    match.click();
    await sleep(250);
    document.body.click(); // cerrar si sigue abierto (multiselect)
    await sleep(150);
    purgeStalePanels(); // no dejar el panel pegado en el DOM para la siguiente llamada
    return true;
  }

  // Busca el input de monto de una fila (Centro de Costos o Concepto) cuyo texto
  // visible coincide con `nombre` (puede llevar sufijo "(Departamentos)").
  // v2.4 (bug real, visto en vivo): nombres CORTOS de Centro de Costos —como
  // "Centro"— coincidían por "empieza con" con textos de la propia pantalla que
  // nada tienen que ver con esa fila (el título "Centro de Costos", la etiqueta del
  // campo "Centro costos"), que además aparecen ANTES en el HTML que la fila real
  // "Centro (Sucursales)". Como el código se quedaba con la PRIMERA coincidencia,
  // terminaba escribiendo el monto en un input equivocado (o en ninguno visible) en
  // vez de en la fila correcta. Ahora primero se busca una coincidencia EXACTA del
  // texto completo de la fila (que si trae sufijo de tipo, como "Centro
  // (Sucursales)", ya no se confunde con nada); "empieza con" queda solo como
  // respaldo para cuando no hay sufijo (ej. filas de Concepto).
  function findRowAmountInput(nombre) {
    const norm = (s) => s.trim().toLowerCase();
    const leaves = Array.from(document.querySelectorAll('div, label, span')).filter(
      (e) => e.children.length === 0
    );
    let candidates = leaves.filter((e) => norm(e.textContent) === norm(nombre));
    if (!candidates.length) {
      candidates = leaves.filter((e) => norm(e.textContent).startsWith(norm(nombre)));
    }
    for (const c of candidates) {
      const row = c.closest('div');
      const input = row ? row.parentElement.querySelector('input') : null;
      if (input) return input;
    }
    return null;
  }

  const MESES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];

  async function setDateField(labelText, fechaStr) {
    const m = fechaStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) throw new Error(`formato de fecha no reconocido: "${fechaStr}"`);
    const day = parseInt(m[1], 10);
    const month = parseInt(m[2], 10); // 1-12
    const year = parseInt(m[3], 10);

    const field = findFieldContainer(labelText);
    if (!field) throw new Error(`No se encontró el campo "${labelText}"`);
    const input = field.querySelector('input');
    input.click();
    await sleep(350);

    const targetValue = year * 12 + (month - 1);

    for (let i = 0; i < 36; i++) {
      const panel = document.querySelector('.p-datepicker');
      if (!panel) throw new Error('no abrió el calendario');
      const title = panel.querySelector('.p-datepicker-title')?.textContent.trim() || '';
      const yearMatch = title.match(/(\d{4})/);
      const currentYear = yearMatch ? parseInt(yearMatch[1], 10) : year;
      const monthName = MESES.find((mm) => title.includes(mm));
      const currentMonthIdx = monthName ? MESES.indexOf(monthName) : month - 1;
      const currentValue = currentYear * 12 + currentMonthIdx;
      if (currentValue === targetValue) break;
      const btn = panel.querySelector(
        currentValue < targetValue ? '.p-datepicker-next' : '.p-datepicker-prev'
      );
      if (!btn) throw new Error('no encontré la navegación del calendario');
      btn.click();
      await sleep(180);
    }

    const panel = document.querySelector('.p-datepicker');
    const cells = Array.from(panel.querySelectorAll('td span'));
    const target = cells.find(
      (c) =>
        c.textContent.trim() === String(day) &&
        !c.closest('td').classList.contains('p-datepicker-other-month')
    );
    if (!target) throw new Error(`no encontré el día ${day} en el calendario`);
    target.click();
    await sleep(150);
  }

  // ---------------------------------------------------------------------
  // Parseo de los datos pegados desde Excel
  // ---------------------------------------------------------------------

  function normalizeHeader(h) {
    return h
      .toLowerCase()
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, '');
  }

  // NUEVO v2.0: ahora acepta encabezados + VARIAS filas de datos (por ejemplo,
  // el Excel completo de solicitudes), no solo encabezado + 1 fila.
  function parseTable(text) {
    const lines = text.trim().split(/\r?\n/).filter((l) => l.trim().length);
    if (lines.length < 2) {
      throw new Error('Pega la fila de encabezados y al menos una fila de datos.');
    }
    const delim = lines[0].includes('\t') ? '\t' : ',';
    const headers = lines[0].split(delim).map((h) => h.trim());
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(delim).map((v) => v.trim());
      const row = {};
      headers.forEach((h, idx) => {
        row[normalizeHeader(h)] = values[idx] ?? '';
      });
      rows.push(row);
    }
    return rows;
  }

  const FIELD_ALIASES = {
    folioQintek: ['Folio Qintek', 'Folio', 'Numero de Orden de Pago', 'No. Orden de Pago', 'Orden de Pago'],
    empresaPagadora: ['Empresa Pagadora'],
    proveedor: ['Proveedor', 'Nombre Proveedor'],
    clabeProveedor: ['CLABE Proveedor', 'CLABE'],
    tipoPago: ['Tipo pago', 'Forma Pago'],
    fechaRegistro: ['Fecha Registro', 'Fecha Factura'],
    semana: ['Semana', 'Semana Gasto', 'Semanas de Gasto'],
    folioFactura: ['Folio Factura'],
    departamento: ['Departamento'],
    centroCostos: ['Centro de Costos'],
    conceptos: ['Conceptos', 'Concepto'],
    iva: ['IVA', 'IVA Factura'],
    moneda: ['Moneda', 'Tipo Moneda'],
    metodoPago: ['Metodo Pago', 'Método Pago'],
    comentarios: ['Comentarios'],
    folioUUID: ['Folio UUID', 'UUID'],
  };

  function buildFieldMap(row) {
    const result = {};
    for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
      for (const alias of aliases) {
        const key = normalizeHeader(alias);
        if (row[key] !== undefined && row[key] !== '') {
          result[field] = row[key];
          break;
        }
      }
    }
    return result;
  }

  // NUEVO v2.0: normaliza un folio para compararlos aunque tengan ceros a la
  // izquierda, separadores de miles, espacios, etc. ("0097006" == "97,006" == "97006").
  function normalizeFolio(v) {
    return (v || '').toString().replace(/\D/g, '').replace(/^0+/, '');
  }

  function findRowByFolio(rows, folioBuscado) {
    const target = normalizeFolio(folioBuscado);
    if (!target) return null;
    return (
      rows.find((row) => {
        const map = buildFieldMap(row);
        return normalizeFolio(map.folioQintek) === target;
      }) || null
    );
  }

  // "2-Administración y Finanzas: $1101.2, 3-Andares: $4225.06" -> [{nombre, monto}]
  function parseNameAmountList(str) {
    if (!str) return [];
    return str
      .split(',')
      .map((part) => {
        const m = part.trim().match(/^\d+-(.+?):\s*\$?([\d.]+)\s*$/);
        if (!m) return null;
        return { nombre: m[1].trim(), monto: parseFloat(m[2]) };
      })
      .filter(Boolean);
  }

  const stripPrefix = (v) => (v || '').replace(/^\d+-/, '').trim();

  // ---------------------------------------------------------------------
  // Flujo principal: completar SOLO lo que falte tras el XML de Qintek
  // ---------------------------------------------------------------------

  async function autocompletar(data) {
    const log = [];
    const ok = (msg) => log.push({ level: 'ok', msg });
    const skip = (msg) => log.push({ level: 'skip', msg });
    const warn = (msg) => log.push({ level: 'warn', msg });
    const err = (msg) => log.push({ level: 'err', msg });

    // 1. Empresa Pagadora — el XML de la factura no trae esto (es la empresa del
    //    grupo que paga, no algo que venga en un CFDI), así que el script siempre
    //    intenta llenarlo con el Excel.
    if (data.empresaPagadora) {
      try {
        const nombre = stripPrefix(data.empresaPagadora);
        await selectDropdownOption('Empresa Pagadora', nombre);
        ok(`Empresa Pagadora: ${nombre}`);
      } catch (e) {
        err(`Empresa Pagadora: ${e.message}`);
      }
    } else {
      warn('Empresa Pagadora: no vino en los datos pegados, selecciónala manualmente.');
    }
    await sleep(300);

    // 2. Proveedor — normalmente ya lo llenó Qintek al leer el XML. Solo se toca
    //    si sigue vacío.
    try {
      if (!isDropdownEmpty('Proveedor')) {
        skip('Proveedor: ya lo llenó Qintek con el XML (sin cambios).');
      } else if (data.proveedor) {
        await selectDropdownOption('Proveedor', data.proveedor);
        ok(`Proveedor: ${data.proveedor}`);
      } else {
        warn(
          `Proveedor: sigue vacío y no vino el nombre en el Excel (solo CLABE ` +
            `${data.clabeProveedor || 'N/D'}). El catálogo de Qintek no expone la CLABE ` +
            'en pantalla, así que no se puede resolver automáticamente. Selecciónalo manualmente.'
        );
      }
    } catch (e) {
      err(`Proveedor: ${e.message}`);
    }
    await sleep(300);

    // 3. Cuenta Destino (Proveedor) — igual que Proveedor, normalmente ya viene del XML.
    try {
      if (!isDropdownEmpty('Cuenta Destino (Proveedor)')) {
        skip('Cuenta Destino: ya la llenó Qintek con el XML (sin cambios).');
      } else if (data.proveedor) {
        const panel = await openFieldPanel('Cuenta Destino (Proveedor)');
        if (!panel) throw new Error('no abrió el panel');
        const items = Array.from(panel.querySelectorAll('li'));
        let match = null;
        if (items.length === 1) match = items[0];
        else if (data.clabeProveedor)
          match = items.find((li) => data.clabeProveedor.includes(li.textContent.trim()));
        if (match) {
          match.click();
          ok(`Cuenta Destino: ${match.textContent.trim()}`);
        } else {
          document.body.click();
          warn(
            `Cuenta Destino: hay varias cuentas para este proveedor, selecciona la que ` +
              `corresponda a la CLABE ${data.clabeProveedor || ''}.`
          );
        }
      }
    } catch (e) {
      err(`Cuenta Destino: ${e.message}`);
    }
    await sleep(300);

    // 4. Concepto(s) + montos — no viene en el XML de la factura, siempre se llena
    //    con el Excel.
    const conceptoItems = parseNameAmountList(data.conceptos);
    for (const item of conceptoItems) {
      try {
        await selectDropdownOption('Concepto', item.nombre);
        ok(`Concepto agregado: ${item.nombre}`);
      } catch (e) {
        err(`Concepto "${item.nombre}": ${e.message}`);
      }
    }
    await sleep(300);
    for (const item of conceptoItems) {
      const input = findRowAmountInput(item.nombre);
      if (input) {
        setAngularInputValue(input, String(item.monto));
        ok(`Monto Concepto "${item.nombre}": $${item.monto}`);
      } else {
        warn(`No encontré el campo de monto para el concepto "${item.nombre}".`);
      }
    }

    // 5. Método Pago (default PUE si no se especifica) — siempre lo llena el script.
    try {
      const valor = data.metodoPago || 'PUE - Pago en Una sola Exhibición';
      await selectDropdownOption('Método Pago', valor);
      ok(`Método Pago: ${valor}`);
    } catch (e) {
      err(`Método Pago: ${e.message}`);
    }

    // 6. Forma Pago (viene de "Tipo pago" en el export) — siempre lo llena el script.
    if (data.tipoPago) {
      try {
        const nombre = stripPrefix(data.tipoPago);
        await selectDropdownOption('Forma Pago', nombre);
        ok(`Forma Pago: ${nombre}`);
      } catch (e) {
        err(`Forma Pago: ${e.message}`);
      }
    } else {
      warn('Forma Pago: no vino en los datos, selecciónala manualmente.');
    }

    // 7. Semana Gasto — siempre lo llena el script.
    const semanaNum = (data.semana || '').split(':')[0].trim();
    if (semanaNum) {
      try {
        await selectDropdownOption('Semana Gasto', semanaNum);
        ok(`Semana Gasto: ${semanaNum}`);
      } catch (e) {
        err(`Semana Gasto: ${e.message}`);
      }
    }
    await sleep(300);

    // 8. Centro de Costos — no viene en el XML de la factura, siempre se llena con
    //    el Excel. IMPORTANTE: "Tipo Centro costos" es en realidad un MULTISELECT, no
    //    un dropdown de una sola opción. Eso trajo varios bugs seguidos (ver CAMBIOS
    //    v1.1, v1.2, v2.2, v2.3, v2.4 arriba). El más reciente y más importante (v2.5):
    //    CADA VEZ que el script cambia de un tipo a otro —no solo al llegar a cero
    //    tipos marcados— Qintek borra los Centro de Costos ya agregados que pertenecían
    //    al tipo que se está desmarcando. Por eso ahora, en vez de ir cambiando de tipo
    //    sobre la marcha según aparecen las líneas, primero se resuelve el tipo de CADA
    //    línea y se arma el conjunto de tipos que realmente se van a necesitar; si son 1
    //    o 2, se marcan TODOS de una vez al principio y ya no se vuelven a tocar.
    const TIPOS_CC = ['Departamentos', 'Sucursales', 'Planta'];
    const ccItems = parseNameAmountList(data.centroCostos);
    const resueltos = [];

    if (ccItems.length) {
      // 8a. Construir un catálogo { tipo: [nombres disponibles] } visitando cada tipo una vez.
      const catalogoPorTipo = {};
      for (const tipo of TIPOS_CC) {
        try {
          await selectDropdownOption('Tipo Centro costos', tipo); // marcar
          await sleep(300);
          const panel = await openFieldPanel('Centro costos');
          if (panel) await waitForStableOptions(panel);
          const opciones = panel
            ? Array.from(panel.querySelectorAll('li')).map((li) => li.textContent.trim())
            : [];
          document.body.click();
          purgeStalePanels();
          catalogoPorTipo[tipo] = opciones;
          await sleep(200);
          await selectDropdownOption('Tipo Centro costos', tipo); // desmarcar (toggle)
        } catch (e) {
          catalogoPorTipo[tipo] = [];
          err(`No pude leer el catálogo de Centro de Costos para "${tipo}": ${e.message}`);
        }
        await sleep(200);
      }

      // 8b. Resolver a qué tipo pertenece cada línea ANTES de tocar nada en el
      //     formulario (esto es lo nuevo en v2.5 — antes se resolvía y se marcaba al
      //     mismo tiempo, línea por línea, lo que forzaba cambios de tipo a media
      //     ejecución cada vez que el Excel alternaba entre tipos).
      // v2.8 (bug real, visto en vivo — la regla de "preferir Planta" de v2.7 podía
      // arruinar TODO el lote): si un lote solo necesitaba 2 tipos (ej. Departamentos +
      // Sucursales, el caso seguro) y una línea ambigua se resolvía a Planta a la
      // fuerza, el lote pasaba a necesitar 3 tipos — y ahí sí aplica la limitación
      // grave de Qintek (cada cambio de tipo borra lo ya agregado del tipo que se
      // desmarca), arriesgando TODAS las líneas, no solo la ambigua. Ahora la
      // resolución se hace en dos pasadas: primero las líneas sin ambigüedad (que
      // determinan qué tipos ya son necesarios de todos modos), y después las
      // ambiguas, evitando siempre que sea posible abrir un tercer tipo.
      const itemsConTipo = [];
      const ambiguos = [];
      for (const item of ccItems) {
        const tiposConMatch = TIPOS_CC.filter((tipo) =>
          (catalogoPorTipo[tipo] || []).some(
            (opt) => opt.toLowerCase() === `${item.nombre} (${tipo})`.toLowerCase()
          )
        );
        if (tiposConMatch.length === 1) {
          itemsConTipo.push({ ...item, tipo: tiposConMatch[0] });
        } else if (tiposConMatch.length > 1) {
          ambiguos.push({ item, tiposConMatch });
        } else {
          warn(`No encontré "${item.nombre}" en Departamentos, Sucursales ni Planta.`);
        }
      }

      const tiposBase = new Set(itemsConTipo.map((i) => i.tipo));
      for (const { item, tiposConMatch } of ambiguos) {
        // Regla acordada con el usuario (v2.7): si "Planta" es una opción, preferirla.
        // Pero (v2.8) solo cuando NO implique abrir un tercer tipo en el lote: si ya se
        // necesitan 2 tipos distintos que no incluyen Planta, forzar Planta aquí
        // pondría en riesgo TODAS las líneas del lote (no solo esta), así que en ese
        // caso se usa en su lugar el otro tipo de la ambigüedad si ya era necesario, o
        // se deja para agregar a mano si ni siquiera eso evita abrir un tercer tipo.
        const abrirPlantaEsSeguro = tiposBase.has('Planta') || tiposBase.size < 2;
        if (tiposConMatch.includes('Planta') && abrirPlantaEsSeguro) {
          itemsConTipo.push({ ...item, tipo: 'Planta' });
          tiposBase.add('Planta');
          warn(
            `Centro de Costos "${item.nombre}" existe en más de un tipo ` +
              `(${tiposConMatch.join(', ')}) — se usó "Planta" por regla acordada.`
          );
          continue;
        }
        const alternativaSegura = tiposConMatch.find((t) => tiposBase.has(t));
        if (alternativaSegura) {
          itemsConTipo.push({ ...item, tipo: alternativaSegura });
          warn(
            `Centro de Costos "${item.nombre}" existe en más de un tipo ` +
              `(${tiposConMatch.join(', ')}) — se usó "${alternativaSegura}" en vez de ` +
              `Planta, para no abrir un tercer tipo en este lote (arriesgaría el resto ` +
              `de las líneas). Verifica si en este caso debía ser Planta.`
          );
        } else {
          warn(
            `Centro de Costos "${item.nombre}" existe en más de un tipo ` +
              `(${tiposConMatch.join(', ')}). Agrégalo manualmente (usar cualquiera de ` +
              `estas opciones aquí abriría un tercer tipo y pondría en riesgo el resto ` +
              `del lote).`
          );
        }
      }

      const tiposNecesarios = [...new Set(itemsConTipo.map((i) => i.tipo))];

      // v2.6 (bug real, visto en vivo — se perdían los montos de las primeras líneas
      // agregadas cuando había muchas): antes se agregaban TODAS las líneas de Centro
      // de Costos primero y hasta el final se recorría `resueltos` para llenar los
      // montos uno por uno. Con órdenes de muchas líneas (11+), los montos de las
      // primeras que se habían agregado ya no se encontraban ("no encontré el campo de
      // monto") — lo más probable es que la lista de filas en pantalla no mantenga
      // renderizadas todas las filas ya agregadas a la vez. Ahora cada línea se agrega
      // Y se le llena el monto de inmediato, uno tras otro, mientras su fila está
      // garantizada en el DOM recién creada.
      const agregarItem = async (item) => {
        try {
          const panel = await openFieldPanel('Centro costos');
          if (!panel) throw new Error('no abrió el panel');
          await waitForStableOptions(panel);
          const options = Array.from(panel.querySelectorAll('li'));
          const norm = item.nombre.toLowerCase();
          const match = options.find(
            (li) => li.textContent.trim().toLowerCase() === norm + ` (${item.tipo})`.toLowerCase()
          );
          if (match) {
            match.click();
            // v2.4: se guarda también el texto EXACTO de la fila (con sufijo de tipo,
            // ej. "Centro (Sucursales)") para poder ubicar su input de monto sin
            // ambigüedad más adelante.
            const etiqueta = match.textContent.trim();
            document.body.click();
            purgeStalePanels();
            await sleep(250);
            const input = findRowAmountInput(etiqueta);
            if (input) {
              setAngularInputValue(input, String(item.monto));
              resueltos.push({ ...item, etiqueta });
              ok(`Centro de Costos agregado: ${etiqueta} — $${item.monto}`);
            } else {
              resueltos.push({ ...item, etiqueta });
              warn(`Centro de Costos "${etiqueta}" se agregó pero no encontré su campo de monto.`);
            }
          } else {
            warn(`No volví a encontrar "${item.nombre}" al reabrir el catálogo de ${item.tipo}.`);
            document.body.click();
            purgeStalePanels();
          }
        } catch (e) {
          err(`Centro de Costos "${item.nombre}": ${e.message}`);
        }
        await sleep(200);
      };

      if (tiposNecesarios.length <= 2) {
        // Caso normal (aunque sean muchas líneas alternando entre los 2 tipos): se
        // marcan de una sola vez los tipos que se van a necesitar y ya NO se vuelven a
        // tocar durante el resto del proceso, así nunca se dispara el borrado de Qintek
        // por cambio de tipo.
        for (const tipo of tiposNecesarios) {
          await selectDropdownOption('Tipo Centro costos', tipo);
          await sleep(300);
        }
        for (const item of itemsConTipo) {
          await agregarItem(item);
        }
      } else {
        // Caso raro: se necesitan los 3 tipos a la vez. Marcarlos los 3 juntos
        // dispara el comportamiento especial "-TODOS-" de Qintek (selecciona solo
        // todos los centros de costo del catálogo), así que aquí sí hay que ir
        // cambiando de tipo sobre la marcha — con la limitación conocida de que el
        // primer tipo procesado puede perderse al llegar al último (ver LÍMITES
        // CONOCIDOS). Se avisa explícitamente antes de empezar.
        warn(
          'Esta orden necesita los 3 tipos de Centro de Costos a la vez ' +
            '(Departamentos + Sucursales + Planta). Por una limitación conocida de ' +
            'Qintek, es posible que se pierdan los del primer tipo procesado — revisa ' +
            'el formulario completo antes de guardar.'
        );
        let tipoActual = null;
        for (const item of itemsConTipo) {
          if (tipoActual !== item.tipo) {
            try {
              await selectDropdownOption('Tipo Centro costos', item.tipo); // marcar el nuevo primero
              await sleep(300);
              if (tipoActual) {
                await selectDropdownOption('Tipo Centro costos', tipoActual); // desmarcar el anterior después
                await sleep(300);
              }
              tipoActual = item.tipo;
            } catch (e) {
              err(`Centro de Costos "${item.nombre}": ${e.message}`);
              continue;
            }
          }
          await agregarItem(item);
        }
      }
    }
    // A propósito NO se desmarca "Tipo Centro costos" aquí — hacerlo borraría todo lo
    // que se acaba de seleccionar en "Centro costos". Se dejan marcados los tipos
    // usados, que es el estado correcto para ver el resultado.
    document.body.click();
    purgeStalePanels();
    await sleep(300);

    // v2.6: el monto de cada Centro de Costos ya se llenó en el momento (ver
    // `agregarItem` arriba) — ya no se rellena aquí al final por lote.

    // 9. Folio de Factura — normalmente ya lo llenó Qintek con el XML.
    if (!isPlainInputEmpty('Folio de Factura')) {
      skip('Folio de Factura: ya lo llenó Qintek con el XML (sin cambios).');
    } else if (data.folioFactura) {
      const input = getFieldInput('Folio de Factura');
      if (input) {
        setAngularInputValue(input, data.folioFactura);
        ok(`Folio de Factura: ${data.folioFactura}`);
      }
    }

    // 10. Fecha factura — normalmente ya la llenó Qintek con el XML.
    if (!isPlainInputEmpty('Fecha factura')) {
      skip('Fecha factura: ya la llenó Qintek con el XML (sin cambios).');
    } else if (data.fechaRegistro) {
      try {
        await setDateField('Fecha factura', data.fechaRegistro);
        ok('Fecha factura capturada.');
      } catch (e) {
        err(`Fecha factura: ${e.message}`);
      }
    }

    // 11. Confirmar Monto Factura (incluye IVA) — normalmente ya lo llenó Qintek con
    //     el XML. Como respaldo, se usa la suma de Centro de Costos (o Conceptos si
    //     no hubo desglose por Centro de Costos).
    if (!isPlainInputEmpty('Confirmar Monto Factura (incluye IVA)')) {
      skip('Confirmar Monto Factura: ya lo llenó Qintek con el XML (sin cambios).');
    } else {
      const total =
        resueltos.reduce((s, i) => s + i.monto, 0) ||
        conceptoItems.reduce((s, i) => s + i.monto, 0);
      if (total) {
        const input = getFieldInput('Confirmar Monto Factura (incluye IVA)');
        if (input) {
          setAngularInputValue(input, String(total));
          ok(`Confirmar Monto Factura: $${total}`);
        }
      }
    }

    // 12. IVA Factura — normalmente ya lo llenó Qintek con el XML.
    try {
      if (!isDropdownEmpty('IVA Factura')) {
        skip('IVA Factura: ya lo llenó Qintek con el XML (sin cambios).');
      } else {
        const valor = data.iva || '16%';
        await selectDropdownOption('IVA Factura', valor);
        ok(`IVA Factura: ${valor}`);
      }
    } catch (e) {
      err(`IVA Factura: ${e.message}`);
    }

    // 13. Tipo Moneda — normalmente ya lo llenó Qintek con el XML.
    try {
      if (!isDropdownEmpty('Tipo Moneda')) {
        skip('Tipo Moneda: ya lo llenó Qintek con el XML (sin cambios).');
      } else {
        const valor = data.moneda || 'MXN';
        await selectDropdownOption('Tipo Moneda', valor);
        ok(`Tipo Moneda: ${valor}`);
      }
    } catch (e) {
      err(`Tipo Moneda: ${e.message}`);
    }

    // 14. Comentarios — no viene en el XML de la factura, siempre se llena con el Excel.
    if (data.comentarios) {
      const field = findFieldContainer('Comentarios');
      const input = field ? field.querySelector('input, textarea') : null;
      if (input) {
        setAngularInputValue(input, data.comentarios);
        ok('Comentarios agregados.');
      }
    }

    // 15. Folio UUID — normalmente ya lo llenó Qintek con el XML.
    if (!isPlainInputEmpty('Folio UUID')) {
      skip('Folio UUID: ya lo llenó Qintek con el XML (sin cambios).');
    } else if (data.folioUUID) {
      const input = getFieldInput('Folio UUID');
      if (input) {
        setAngularInputValue(input, data.folioUUID);
        ok('Folio UUID agregado.');
      }
    }

    return log;
  }

  // ---------------------------------------------------------------------
  // Panel flotante (interfaz)
  // ---------------------------------------------------------------------

  function renderLog(container, log) {
    const colors = { ok: '#2e7d32', skip: '#1565c0', warn: '#e65100', err: '#c62828', info: '#555' };
    const icons = { ok: '✅', skip: '⏭️', warn: '⚠️', err: '❌', info: '🔎' };
    container.innerHTML = log
      .map(
        (l) =>
          `<div style="color:${colors[l.level]}; margin-bottom:3px;">${icons[l.level]} ${l.msg}</div>`
      )
      .join('');
  }

  function createPanel() {
    if (document.getElementById('qintek-autofill-panel')) return;

    const panel = document.createElement('div');
    panel.id = 'qintek-autofill-panel';
    panel.style.cssText = `
      position: fixed; top: 80px; right: 20px; width: 380px; z-index: 999999;
      background: #fff; border: 2px solid #f57c00; border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.25); font-family: Arial, sans-serif; font-size: 13px;
    `;
    panel.innerHTML = `
      <div style="background:#f57c00; color:#fff; padding:8px 12px; border-radius:6px 6px 0 0; display:flex; justify-content:space-between; align-items:center; cursor:default;">
        <strong>Panel de Clonado de Orden de Pago</strong>
        <span id="qaf-toggle" style="cursor:pointer; padding:0 4px;">▁</span>
      </div>
      <div id="qaf-body" style="padding:12px;">
        <div style="margin-bottom:6px; color:#555;">
          1) Sube primero el XML de la factura en Qintek. 2) Pega aquí tu Excel completo
          (encabezados + todas las filas de órdenes que podrías clonar).
        </div>
        <textarea id="qaf-input" rows="4" style="width:100%; box-sizing:border-box; font-family:monospace; font-size:11px;"
          placeholder="Folio Qintek	Empresa Pagadora	...&#10;0097006	1-Laguneros Orientales...	...&#10;0098012	2-Otra Empresa...	..."></textarea>
        <div style="margin:8px 0;">
          <label style="display:block; margin-bottom:4px;">Número de orden de pago (folio) a clonar:</label>
          <input id="qaf-folio" type="text" style="width:100%; box-sizing:border-box;" placeholder="Ej. 97006 (opcional si solo pegaste una fila)" />
        </div>
        <div style="margin:8px 0;">
          <label style="display:block; margin-bottom:4px;">
            Proveedor (solo si sigue vacío tras el XML y el Excel no lo trae — escribe el nombre exacto de Qintek):
          </label>
          <input id="qaf-proveedor" type="text" style="width:100%; box-sizing:border-box;" placeholder="Opcional" />
        </div>
        <button id="qaf-run" style="width:100%; padding:8px; background:#f57c00; color:#fff; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">
          Buscar folio y completar faltantes
        </button>
        <div id="qaf-log" style="margin-top:10px; max-height:220px; overflow:auto;"></div>
        <div style="margin-top:8px; font-size:11px; color:#b00020; font-weight:bold;">
          ⚠ Este script NUNCA presiona "Guardar". Revisa todo y guarda tú mismo.
        </div>
      </div>
    `;
    document.body.appendChild(panel);

    document.getElementById('qaf-toggle').addEventListener('click', () => {
      const body = document.getElementById('qaf-body');
      body.style.display = body.style.display === 'none' ? 'block' : 'none';
    });

    document.getElementById('qaf-run').addEventListener('click', async () => {
      const runBtn = document.getElementById('qaf-run');
      const logDiv = document.getElementById('qaf-log');
      const pre = [];
      runBtn.disabled = true;
      runBtn.textContent = 'Buscando...';
      try {
        const raw = document.getElementById('qaf-input').value;
        const folioInput = document.getElementById('qaf-folio').value.trim();
        const rows = parseTable(raw);

        let row;
        if (rows.length === 1) {
          row = rows[0];
          if (folioInput) {
            const map = buildFieldMap(row);
            if (normalizeFolio(map.folioQintek) && normalizeFolio(map.folioQintek) !== normalizeFolio(folioInput)) {
              pre.push({
                level: 'warn',
                msg: `El folio de la única fila pegada (${map.folioQintek}) no coincide con "${folioInput}"; se usó de todos modos por ser la única fila.`,
              });
            }
          }
        } else {
          if (!folioInput) {
            throw new Error(
              `Pegaste ${rows.length} filas: escribe el "Número de orden de pago" (folio) para elegir cuál usar.`
            );
          }
          row = findRowByFolio(rows, folioInput);
          if (!row) {
            throw new Error(`No encontré el folio "${folioInput}" entre las ${rows.length} filas pegadas.`);
          }
        }

        const data = buildFieldMap(row);
        pre.push({
          level: 'info',
          msg: `Usando la orden con folio ${data.folioQintek || folioInput || '(sin folio)'} — ${stripPrefix(data.empresaPagadora || '') || 'empresa no especificada'}.`,
        });

        const proveedorManual = document.getElementById('qaf-proveedor').value.trim();
        if (proveedorManual) data.proveedor = proveedorManual;

        runBtn.textContent = 'Completando...';
        const log = await autocompletar(data);
        renderLog(logDiv, pre.concat(log));
      } catch (e) {
        renderLog(logDiv, pre.concat([{ level: 'err', msg: e.message }]));
      } finally {
        runBtn.disabled = false;
        runBtn.textContent = 'Buscar folio y completar faltantes';
      }
    });
  }

  // Espera a que el formulario Angular esté renderizado antes de inyectar el panel.
  // NUEVO v2.1: ya NO se detiene tras la primera vez. En una app Angular como Qintek,
  // navegar entre pantallas del menú lateral no siempre recarga la página completa
  // (routing del lado del cliente), así que el panel podía desaparecer sin que
  // Tampermonkey volviera a inyectar el script. createPanel() ya es seguro de llamar
  // repetidamente (no hace nada si el panel ya existe), así que este chequeo cada
  // segundo actúa como "auto-recuperación": si el panel se pierde, se vuelve a crear
  // en cuanto la pantalla de Órdenes de pago esté visible de nuevo.
  setInterval(() => {
    if (document.querySelector('label')) {
      createPanel();
    }
  }, 1000);

  // NUEVO v2.1: forma manual de "mandar traer" el panel en cualquier momento, sin
  // depender de que se haya inyectado solo ni de refrescar la página. Aparece en el
  // menú del ícono de Tampermonkey (en la barra de Chrome) mientras estés en una
  // pestaña de Qintek que coincida con esta URL.
  if (typeof GM_registerMenuCommand === 'function') {
    GM_registerMenuCommand('Mostrar Panel de Clonado', () => {
      createPanel();
      const panelEl = document.getElementById('qintek-autofill-panel');
      if (panelEl) {
        const body = document.getElementById('qaf-body');
        if (body) body.style.display = 'block';
        panelEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    });
  }
})();
