let datosOriginales = [];
let datosFiltrados = [];

let graficoClientes = null;
let graficoLineas = null;
let graficoMensual = null;
let graficoVendedores = null;


// ======================================================
// CARGAR ARCHIVO EXCEL
// ======================================================

document
    .getElementById("archivoExcel")
    .addEventListener("change", cargarExcel);


function cargarExcel(event) {

    const archivo = event.target.files[0];

    const estado = document.getElementById("estadoArchivo");

    if (!archivo) {
        estado.textContent = "Ningún archivo cargado";
        return;
    }

    estado.textContent = "Leyendo archivo...";

    const lector = new FileReader();

    lector.onload = function(e) {

        try {

            const datos = new Uint8Array(e.target.result);

            const libro = XLSX.read(datos, {
                type: "array"
            });

            if (!libro.SheetNames.length) {
                throw new Error("El archivo no contiene hojas.");
            }

            const nombreHoja = libro.SheetNames[0];

            const hoja = libro.Sheets[nombreHoja];

            const filas = XLSX.utils.sheet_to_json(hoja, {
                defval: ""
            });

            if (!filas.length) {
                throw new Error(
                    "La primera hoja del Excel no contiene información."
                );
            }

            console.log("Columnas encontradas:", Object.keys(filas[0]));
            console.log("Primera fila:", filas[0]);

            datosOriginales = prepararDatos(filas);

            if (!datosOriginales.length) {
                throw new Error(
                    "No se pudieron preparar los datos del Excel."
                );
            }

            llenarFiltros();

            actualizarDashboard();

            estado.textContent =
                "Archivo cargado correctamente: " +
                datosOriginales.length +
                " registros.";

        }

        catch (error) {

            console.error(
                "ERROR COMPLETO:",
                error
            );

            estado.textContent =
                "ERROR: " +
                error.message;

        }

    };

    lector.onerror = function() {

        estado.textContent =
            "No fue posible leer el archivo.";

    };

    lector.readAsArrayBuffer(archivo);
}


// ======================================================
// PREPARAR DATOS
// ======================================================

function prepararDatos(filas) {

    return filas.map(function(fila) {

        const ventas = limpiarNumero(
            fila["VrRem"] ??
            fila["Vr Rem"] ??
            fila["VRREM"] ??
            fila["Valor Remisión"] ??
            fila["Valor Remision"] ??
            0
        );

        const costo = limpiarNumero(
            fila["Costo"] ??
            fila["COSTO"] ??
            fila["Costos"] ??
            0
        );

        const utilidadExcel = limpiarNumero(
            fila["Utilidad"] ?? 0
        );

        const utilidad =
            ventas - costo;

        const margen =
            ventas !== 0
                ? utilidad / ventas
                : 0;

        return {

            vendedor:
                limpiarTexto(
                    fila["Vendedor"]
                ),

            cliente:
                limpiarTexto(
                    fila["Cliente"]
                ),

            remision:
                limpiarTexto(
                    fila["Remision"]
                ),

            fecha:
                fila["Fecha"],

            pedido:
                limpiarTexto(
                    fila["Pedid"]
                ),

            linea:
                limpiarTexto(
                    fila["Linea"]
                ),

            codigo:
                limpiarTexto(
                    fila["Cod"]
                ),

            llave:
                limpiarTexto(
                    fila["Llave1"]
                ),

            producto:
                limpiarTexto(
                    fila["Producto"]
                ),

            cantidad:
                limpiarNumero(
                    fila["Cant"]
                ),

            mes:
                obtenerMes(
                    fila["Mes"],
                    fila["Fecha"]
                ),

            factura:
                limpiarTexto(
                    fila["Factura"]
                ),

            valorVenta:
                ventas,

            costo:
                costo,

            utilidad:
                utilidad,

            utilidadExcel:
                utilidadExcel,

            margen:
                margen,

            anio:
                obtenerAnio(
                    fila["Año"],
                    fila["Fecha"]
                )

        };

    });

}


// ======================================================
// LIMPIAR TEXTO
// ======================================================

function limpiarTexto(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {
        return "";
    }

    return String(valor).trim();

}


// ======================================================
// LIMPIAR NÚMEROS
// ======================================================

function limpiarNumero(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return 0;
    }

    if (typeof valor === "number") {
        return Number.isFinite(valor)
            ? valor
            : 0;
    }

    let texto =
        String(valor)
        .trim()
        .replace(/\s/g, "");

    if (!texto) {
        return 0;
    }

    // Elimina símbolos de moneda
    texto = texto.replace(
        /[$€£]/g,
        ""
    );

    // Caso colombiano:
    // 9.460.000
    // 3.691.869
    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {

        texto =
            texto
            .replace(/\./g, "")
            .replace(",", ".");

    }

    // Caso:
    // 9.460.000
    if (
        texto.includes(".") &&
        !texto.includes(",")
    ) {

        const partes =
            texto.split(".");

        const parecenMiles =
            partes.length > 2 ||
            (
                partes.length === 2 &&
                partes[1].length === 3
            );

        if (parecenMiles) {

            texto =
                texto.replace(/\./g, "");

        }

    }

    // Caso:
    // 9460000,50
    if (
        texto.includes(",") &&
        !texto.includes(".")
    ) {

        const partes =
            texto.split(",");

        if (
            partes.length === 2 &&
            partes[1].length <= 2
        ) {

            texto =
                texto.replace(",", ".");

        } else {

            texto =
                texto.replace(/,/g, "");

        }

    }

    // Dejar solamente números,
    // punto y signo negativo
    texto =
        texto.replace(
            /[^0-9.-]/g,
            ""
        );

    const numero =
        Number(texto);

    return Number.isFinite(numero)
        ? numero
        : 0;

}


// ======================================================
// OBTENER AÑO
// ======================================================

function obtenerAnio(valor, fecha) {

    if (
        valor !== undefined &&
        valor !== null &&
        valor !== ""
    ) {

        const numero =
            Number(valor);

        if (
            Number.isFinite(numero) &&
            numero > 1900
        ) {
            return numero;
        }

    }

    if (
        fecha instanceof Date &&
        !isNaN(fecha)
    ) {

        return fecha.getFullYear();

    }

    if (typeof fecha === "number") {

        const fechaExcel =
            XLSX.SSF.parse_date_code(
                fecha
            );

        if (fechaExcel) {
            return fechaExcel.y;
        }

    }

    const texto =
        String(fecha || "");

    const coincidencia =
        texto.match(
            /(20\d{2})/
        );

    if (coincidencia) {
        return Number(
            coincidencia[1]
        );
    }

    return 2026;

}


// ======================================================
// OBTENER MES
// ======================================================

function obtenerMes(valor, fecha) {

    if (
        valor !== undefined &&
        valor !== null &&
        valor !== ""
    ) {

        const numero =
            Number(valor);

        if (
            Number.isFinite(numero) &&
            numero >= 1 &&
            numero <= 12
        ) {
            return numero;
        }

    }

    if (
        fecha instanceof Date &&
        !isNaN(fecha)
    ) {

        return fecha.getMonth() + 1;

    }

    if (typeof fecha === "number") {

        const fechaExcel =
            XLSX.SSF.parse_date_code(
                fecha
            );

        if (fechaExcel) {
            return fechaExcel.m;
        }

    }

    const texto =
        String(fecha || "");

    const coincidencia =
        texto.match(
            /(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})/
        );

    if (coincidencia) {

        return Number(
            coincidencia[2]
        );

    }

    return 1;

}


// ======================================================
// LLENAR FILTROS
// ======================================================

function llenarFiltros() {

    llenarSelect(
        "filtroAnio",
        datosOriginales.map(
            d => d.anio
        ),
        "Todos"
    );

    llenarSelect(
        "filtroMes",
        datosOriginales.map(
            d => d.mes
        ),
        "Todos"
    );

    llenarSelect(
        "filtroVendedor",
        datosOriginales.map(
            d => d.vendedor
        ),
        "Todos"
    );

    llenarSelect(
        "filtroCliente",
        datosOriginales.map(
            d => d.cliente
        ),
        "Todos"
    );

    llenarSelect(
        "filtroLinea",
        datosOriginales.map(
            d => d.linea
        ),
        "Todas"
    );

    llenarSelect(
        "filtroProducto",
        datosOriginales.map(
            d => d.producto
        ),
        "Todos"
    );

}


// ======================================================
// CREAR OPCIONES DE SELECT
// ======================================================

function llenarSelect(
    id,
    valores,
    textoTodos
) {

    const select =
        document.getElementById(id);

    if (!select) {
        return;
    }

    select.innerHTML = "";

    const opcionTodos =
        document.createElement("option");

    opcionTodos.value = "todos";
    opcionTodos.textContent = textoTodos;

    select.appendChild(
        opcionTodos
    );

    const valoresUnicos =
        [...new Set(
            valores
                .filter(
                    v =>
                        v !== null &&
                        v !== undefined &&
                        v !== ""
                )
                .map(
                    v => String(v)
                )
        )]
        .sort(
            (a, b) =>
                a.localeCompare(
                    b,
                    "es",
                    {
                        numeric: true
                    }
                )
        );

    valoresUnicos.forEach(
        function(valor) {

            const opcion =
                document.createElement(
                    "option"
                );

            opcion.value = valor;
            opcion.textContent = valor;

            select.appendChild(
                opcion
            );

        }
    );

}


// ======================================================
// EVENTOS DE FILTROS
// ======================================================

[
    "filtroAnio",
    "filtroMes",
    "filtroVendedor",
    "filtroCliente",
    "filtroLinea",
    "filtroProducto"
]
.forEach(
    function(id) {

        const elemento =
            document.getElementById(id);

        if (elemento) {

            elemento.addEventListener(
                "change",
                actualizarDashboard
            );

        }

    }
);


// ======================================================
// ACTUALIZAR DASHBOARD
// ======================================================

function actualizarDashboard() {

    datosFiltrados =
        datosOriginales.filter(
            function(dato) {

                return coincideFiltro(
                    dato.anio,
                    "filtroAnio"
                )
                &&
                coincideFiltro(
                    dato.mes,
                    "filtroMes"
                )
                &&
                coincideFiltro(
                    dato.vendedor,
                    "filtroVendedor"
                )
                &&
                coincideFiltro(
                    dato.cliente,
                    "filtroCliente"
                )
                &&
                coincideFiltro(
                    dato.linea,
                    "filtroLinea"
                )
                &&
                coincideFiltro(
                    dato.producto,
                    "filtroProducto"
                );

            }
        );

    actualizarKPIs();

    actualizarGraficos();

    actualizarTabla();

}


// ======================================================
// VALIDAR FILTROS
// ======================================================

function coincideFiltro(
    valor,
    idFiltro
) {

    const filtro =
        document.getElementById(
            idFiltro
        );

    if (!filtro) {
        return true;
    }

    if (
        filtro.value === "todos"
    ) {
        return true;
    }

    return String(valor) ===
        String(filtro.value);

}


// ======================================================
// ACTUALIZAR KPIs
// ======================================================

function actualizarKPIs() {

    let ventas = 0;
    let costos = 0;
    let utilidad = 0;

    datosFiltrados.forEach(
        function(dato) {

            ventas +=
                Number(dato.valorVenta) || 0;

            costos +=
                Number(dato.costo) || 0;

            utilidad +=
                Number(dato.utilidad) || 0;

        }
    );

    const margen =
        ventas !== 0
            ? utilidad / ventas
            : 0;

    document.getElementById(
        "ventas"
    ).textContent =
        formatoMoneda(ventas);

    document.getElementById(
        "costos"
    ).textContent =
        formatoMoneda(costos);

    document.getElementById(
        "utilidad"
    ).textContent =
        formatoMoneda(utilidad);

    document.getElementById(
        "margen"
    ).textContent =
        formatoPorcentaje(margen);

}


// ======================================================
// GRÁFICOS
// ======================================================

function actualizarGraficos() {

    crearGraficoClientes();

    crearGraficoLineas();

    crearGraficoMensual();

    crearGraficoVendedores();

}


// ======================================================
// AGRUPAR DATOS
// ======================================================

function agrupar(
    campo,
    tipo = "utilidad"
) {

    const resultado = {};

    datosFiltrados.forEach(
        function(dato) {

            const clave =
                dato[campo] ||
                "Sin información";

            if (!resultado[clave]) {

                resultado[clave] = {
                    ventas: 0,
                    costos: 0,
                    utilidad: 0
                };

            }

            resultado[clave].ventas +=
                Number(
                    dato.valorVenta
                ) || 0;

            resultado[clave].costos +=
                Number(
                    dato.costo
                ) || 0;

            resultado[clave].utilidad +=
                Number(
                    dato.utilidad
                ) || 0;

        }
    );

    return resultado;

}


// ======================================================
// GRÁFICO CLIENTES
// ======================================================

function crearGraficoClientes() {

    const canvas =
        document.getElementById(
            "graficoClientes"
        );

    if (!canvas) {
        return;
    }

    const agrupado =
        agrupar("cliente");

    const elementos =
        Object.entries(agrupado)
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
        )
        .slice(0, 15);

    const etiquetas =
        elementos.map(
            x => x[0]
        );

    const valores =
        elementos.map(
            x => x[1].utilidad
        );

    if (graficoClientes) {
        graficoClientes.destroy();
    }

    graficoClientes =
        new Chart(
            canvas,
            {
                type: "bar",

                data: {

                    labels: etiquetas,

                    datasets: [
                        {
                            label: "Utilidad",

                            data: valores
                        }
                    ]

                },

                options: {

                    responsive: true,

                    plugins: {

                        legend: {
                            display: false
                        }

                    }

                }

            }
        );

}


// ======================================================
// GRÁFICO LÍNEAS
// ======================================================

function crearGraficoLineas() {

    const canvas =
        document.getElementById(
            "graficoLineas"
        );

    if (!canvas) {
        return;
    }

    const agrupado =
        agrupar("linea");

    const elementos =
        Object.entries(agrupado)
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
        );

    const etiquetas =
        elementos.map(
            x => x[0]
        );

    const valores =
        elementos.map(
            x => x[1].utilidad
        );

    if (graficoLineas) {
        graficoLineas.destroy();
    }

    graficoLineas =
        new Chart(
            canvas,
            {
                type: "bar",

                data: {

                    labels: etiquetas,

                    datasets: [
                        {
                            label: "Utilidad",

                            data: valores
                        }
                    ]

                },

                options: {

                    responsive: true,

                    plugins: {

                        legend: {
                            display: false
                        }

                    }

                }

            }
        );

}


// ======================================================
// GRÁFICO MENSUAL
// ======================================================

function crearGraficoMensual() {

    const canvas =
        document.getElementById(
            "graficoMensual"
        );

    if (!canvas) {
        return;
    }

    const agrupado = {};

    for (
        let mes = 1;
        mes <= 12;
        mes++
    ) {

        agrupado[mes] = 0;

    }

    datosFiltrados.forEach(
        function(dato) {

            const mes =
                Number(dato.mes);

            if (
                mes >= 1 &&
                mes <= 12
            ) {

                agrupado[mes] +=
                    Number(
                        dato.utilidad
                    ) || 0;

            }

        }
    );

    const nombresMeses = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre"
    ];

    const valores =
        nombresMeses.map(
            function(_, index) {

                return agrupado[
                    index + 1
                ];

            }
        );

    if (graficoMensual) {
        graficoMensual.destroy();
    }

    graficoMensual =
        new Chart(
            canvas,
            {
                type: "line",

                data: {

                    labels:
                        nombresMeses,

                    datasets: [
                        {
                            label:
                                "Utilidad",

                            data:
                                valores,

                            tension:
                                0.3
                        }
                    ]

                },

                options: {

                    responsive: true

                }

            }
        );

}


// ======================================================
// GRÁFICO VENDEDORES
// ======================================================

function crearGraficoVendedores() {

    const canvas =
        document.getElementById(
            "graficoVendedores"
        );

    if (!canvas) {
        return;
    }

    const agrupado =
        agrupar("vendedor");

    const elementos =
        Object.entries(agrupado)
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
        );

    const etiquetas =
        elementos.map(
            x => x[0]
        );

    const valores =
        elementos.map(
            x => x[1].utilidad
        );

    if (graficoVendedores) {
        graficoVendedores.destroy();
    }

    graficoVendedores =
        new Chart(
            canvas,
            {
                type: "bar",

                data: {

                    labels: etiquetas,

                    datasets: [
                        {
                            label:
                                "Utilidad",

                            data:
                                valores
                        }
                    ]

                },

                options: {

                    indexAxis: "y",

                    responsive: true,

                    plugins: {

                        legend: {
                            display: false
                        }

                    }

                }

            }
        );

}


// ======================================================
// TABLA DE PRODUCTOS
// ======================================================

function actualizarTabla() {

    const cuerpo =
        document.getElementById(
            "tablaProductos"
        );

    if (!cuerpo) {
        return;
    }

    cuerpo.innerHTML = "";

    const agrupado =
        agrupar("producto");

    const productos =
        Object.entries(agrupado)
        .sort(
            (a, b) =>
                b[1].ventas -
                a[1].ventas
        );

    if (!productos.length) {

        cuerpo.innerHTML = `
            <tr>
                <td colspan="5">
                    No hay información
                    para los filtros seleccionados.
                </td>
            </tr>
        `;

        return;

    }

    productos.forEach(
        function([producto, datos]) {

            const margen =
                datos.ventas !== 0
                    ? datos.utilidad /
                      datos.ventas
                    : 0;

            const fila =
                document.createElement(
                    "tr"
                );

            fila.innerHTML = `

                <td>
                    ${escapeHTML(producto)}
                </td>

                <td>
                    ${formatoMoneda(
                        datos.ventas
                    )}
                </td>

                <td>
                    ${formatoMoneda(
                        datos.costos
                    )}
                </td>

                <td>
                    ${formatoMoneda(
                        datos.utilidad
                    )}
                </td>

                <td>
                    ${formatoPorcentaje(
                        margen
                    )}
                </td>

            `;

            cuerpo.appendChild(
                fila
            );

        }
    );

}


// ======================================================
// FORMATO MONEDA
// ======================================================

function formatoMoneda(valor) {

    return new Intl.NumberFormat(
        "es-CO",
        {
            style: "currency",
            currency: "COP",
            maximumFractionDigits: 0
        }
    ).format(
        Number(valor) || 0
    );

}


// ======================================================
// FORMATO PORCENTAJE
// ======================================================

function formatoPorcentaje(valor) {

    return new Intl.NumberFormat(
        "es-CO",
        {
            style: "percent",
            minimumFractionDigits: 1,
            maximumFractionDigits: 1
        }
    ).format(
        Number(valor) || 0
    );

}


// ======================================================
// SEGURIDAD PARA TEXTO HTML
// ======================================================

function escapeHTML(valor) {

    return String(valor)
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );

}
