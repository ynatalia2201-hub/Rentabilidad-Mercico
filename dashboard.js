// ============================================================
// DASHBOARD DE RENTABILIDAD - MERCICO
// ============================================================

let datosOriginales = [];
let datosFiltrados = [];

let graficoClientes = null;
let graficoLineas = null;
let graficoMensual = null;
let graficoVendedores = null;


// ============================================================
// INICIO
// ============================================================

document.addEventListener("DOMContentLoaded", function () {

    const archivo =
        document.getElementById("archivoExcel");

    if (!archivo) {
        console.error(
            "No se encontró el elemento archivoExcel"
        );
        return;
    }

    archivo.addEventListener(
        "change",
        cargarExcel
    );

    activarFiltros();

});


// ============================================================
// CARGAR EXCEL
// ============================================================

function cargarExcel(event) {

    const archivo =
        event.target.files[0];

    const estado =
        document.getElementById(
            "estadoArchivo"
        );

    if (!archivo) {

        estado.textContent =
            "Ningún archivo cargado";

        return;
    }

    estado.textContent =
        "Leyendo archivo Excel...";

    const lector =
        new FileReader();

    lector.onload =
        function (e) {

            try {

                const datos =
                    new Uint8Array(
                        e.target.result
                    );

                const libro =
                    XLSX.read(
                        datos,
                        {
                            type: "array",
                            cellDates: true
                        }
                    );

                if (
                    !libro.SheetNames ||
                    !libro.SheetNames.length
                ) {

                    throw new Error(
                        "El Excel no contiene hojas."
                    );

                }

                const nombreHoja =
                    libro.SheetNames[0];

                const hoja =
                    libro.Sheets[nombreHoja];

                const filas =
                    XLSX.utils.sheet_to_json(
                        hoja,
                        {
                            defval: "",
                            raw: true
                        }
                    );

                if (!filas.length) {

                    throw new Error(
                        "La hoja está vacía."
                    );

                }

                console.log(
                    "HOJA:",
                    nombreHoja
                );

                console.log(
                    "COLUMNAS ORIGINALES:",
                    Object.keys(filas[0])
                );

                console.log(
                    "PRIMERA FILA:",
                    filas[0]
                );

                datosOriginales =
                    prepararDatos(
                        filas
                    );

                if (
                    !datosOriginales.length
                ) {

                    throw new Error(
                        "No se pudieron preparar los registros."
                    );

                }

                console.log(
                    "DATOS PREPARADOS:",
                    datosOriginales
                );

                console.log(
                    "PRIMER REGISTRO PREPARADO:",
                    datosOriginales[0]
                );

                llenarFiltros();

                actualizarDashboard();

                estado.textContent =
                    "Archivo cargado correctamente. " +
                    datosOriginales.length +
                    " registros.";

            }

            catch (error) {

                console.error(
                    "ERROR:",
                    error
                );

                estado.textContent =
                    "ERROR: " +
                    error.message;

            }

        };


    lector.onerror =
        function () {

            estado.textContent =
                "No se pudo leer el archivo.";

        };


    lector.readAsArrayBuffer(
        archivo
    );

}


// ============================================================
// PREPARAR DATOS
// ============================================================

function prepararDatos(filas) {

    return filas.map(
        function (fila) {

            const columnas =
                crearMapaColumnas(
                    fila
                );


            // -----------------------------
            // VENTAS
            // -----------------------------

            const ventas =
                obtenerNumeroPorColumnas(
                    columnas,
                    [
                        "vrrem",
                        "valorremision",
                        "valorremision",
                        "valorventa",
                        "ventas",
                        "venta",
                        "vrventa"
                    ]
                );


            // -----------------------------
            // COSTO
            // -----------------------------

            const costo =
                obtenerNumeroPorColumnas(
                    columnas,
                    [
                        "costo",
                        "costos",
                        "costototal",
                        "costoproducto"
                    ]
                );


            // -----------------------------
            // UTILIDAD
            // -----------------------------

            let utilidad =
                obtenerNumeroPorColumnas(
                    columnas,
                    [
                        "utilidad",
                        "ganancia",
                        "utilidadbruta"
                    ]
                );


            // Si no existe utilidad,
            // la calculamos.

            if (
                utilidad === 0 &&
                (
                    ventas !== 0 ||
                    costo !== 0
                )
            ) {

                utilidad =
                    ventas -
                    costo;

            }


            // -----------------------------
            // MARGEN
            // -----------------------------

            let margen =
                obtenerNumeroPorColumnas(
                    columnas,
                    [
                        "margen",
                        "margenrentabilidad",
                        "rentabilidad"
                    ]
                );


            // Si el margen viene como 61,
            // convertirlo a 0.61

            if (
                Math.abs(margen) > 1
            ) {

                margen =
                    margen / 100;

            }


            // Si no existe margen,
            // calcularlo.

            if (
                margen === 0 &&
                ventas !== 0
            ) {

                margen =
                    utilidad /
                    ventas;

            }


            return {

                vendedor:
                    obtenerTexto(
                        columnas,
                        [
                            "vendedor",
                            "asesor",
                            "comercial"
                        ]
                    ),

                cliente:
                    obtenerTexto(
                        columnas,
                        [
                            "cliente",
                            "nombrecliente"
                        ]
                    ),

                remision:
                    obtenerTexto(
                        columnas,
                        [
                            "remision",
                            "remisión",
                            "rem"
                        ]
                    ),

                fecha:
                    obtenerValor(
                        columnas,
                        [
                            "fecha",
                            "fecharemision"
                        ]
                    ),

                pedido:
                    obtenerTexto(
                        columnas,
                        [
                            "pedid",
                            "pedido",
                            "numpedido"
                        ]
                    ),

                linea:
                    obtenerTexto(
                        columnas,
                        [
                            "linea",
                            "línea"
                        ]
                    ),

                codigo:
                    obtenerTexto(
                        columnas,
                        [
                            "cod",
                            "codigo",
                            "código"
                        ]
                    ),

                llave:
                    obtenerTexto(
                        columnas,
                        [
                            "llave1",
                            "llave"
                        ]
                    ),

                producto:
                    obtenerTexto(
                        columnas,
                        [
                            "producto",
                            "nombreproducto"
                        ]
                    ),

                cantidad:
                    obtenerNumeroPorColumnas(
                        columnas,
                        [
                            "cant",
                            "cantidad"
                        ]
                    ),

                mes:
                    obtenerMes(
                        columnas
                    ),

                factura:
                    obtenerTexto(
                        columnas,
                        [
                            "factura",
                            "numfactura"
                        ]
                    ),

                valorVenta:
                    ventas,

                costo:
                    costo,

                utilidad:
                    utilidad,

                margen:
                    margen,

                anio:
                    obtenerAnio(
                        columnas
                    )

            };

        }
    );

}


// ============================================================
// MAPEAR COLUMNAS
// ============================================================

function crearMapaColumnas(fila) {

    const mapa = {};

    Object.keys(fila)
        .forEach(
            function (nombreOriginal) {

                const nombreNormalizado =
                    normalizarNombre(
                        nombreOriginal
                    );

                mapa[
                    nombreNormalizado
                ] = fila[
                    nombreOriginal
                ];

            }
        );

    return mapa;

}


// ============================================================
// NORMALIZAR NOMBRES
// ============================================================

function normalizarNombre(
    texto
) {

    return String(
        texto || ""
    )
        .toLowerCase()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /[^a-z0-9]/g,
            ""
        );

}


// ============================================================
// OBTENER VALOR
// ============================================================

function obtenerValor(
    columnas,
    posibles
) {

    for (
        let i = 0;
        i < posibles.length;
        i++
    ) {

        const clave =
            normalizarNombre(
                posibles[i]
            );

        if (
            Object.prototype.hasOwnProperty
                .call(
                    columnas,
                    clave
                )
        ) {

            return columnas[
                clave
            ];

        }

    }

    return "";

}


// ============================================================
// OBTENER TEXTO
// ============================================================

function obtenerTexto(
    columnas,
    posibles
) {

    const valor =
        obtenerValor(
            columnas,
            posibles
        );

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";

    }

    return String(
        valor
    ).trim();

}


// ============================================================
// OBTENER NÚMERO
// ============================================================

function obtenerNumeroPorColumnas(
    columnas,
    posibles
) {

    const valor =
        obtenerValor(
            columnas,
            posibles
        );

    return limpiarNumero(
        valor
    );

}


// ============================================================
// LIMPIAR NÚMEROS
// ============================================================

function limpiarNumero(
    valor
) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return 0;

    }


    if (
        typeof valor === "number"
    ) {

        return Number.isFinite(
            valor
        )
            ? valor
            : 0;

    }


    let texto =
        String(valor)
            .trim();


    if (!texto) {
        return 0;
    }


    // Eliminar espacios

    texto =
        texto.replace(
            /\s/g,
            ""
        );


    // Eliminar monedas

    texto =
        texto.replace(
            /[$€£COP]/gi,
            ""
        );


    // ------------------------------------
    // CASO COLOMBIANO
    //
    // 9.460.000
    // 3.691.869
    // ------------------------------------

    if (
        texto.includes(".") &&
        !texto.includes(",")
    ) {

        const partes =
            texto.split(".");


        if (
            partes.length > 2
        ) {

            texto =
                texto.replace(
                    /\./g,
                    ""
                );

        }

        else if (
            partes.length === 2 &&
            partes[1].length === 3
        ) {

            texto =
                texto.replace(
                    /\./g,
                    ""
                );

        }

    }


    // ------------------------------------
    // 9.460.000,50
    // ------------------------------------

    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {

        texto =
            texto
                .replace(
                    /\./g,
                    ""
                )
                .replace(
                    ",",
                    "."
                );

    }


    // ------------------------------------
    // 9460000,50
    // ------------------------------------

    else if (
        texto.includes(",")
    ) {

        const partes =
            texto.split(",");


        if (
            partes.length === 2 &&
            partes[1].length <= 2
        ) {

            texto =
                texto.replace(
                    ",",
                    "."
                );

        }

        else {

            texto =
                texto.replace(
                    /,/g,
                    ""
                );

        }

    }


    // Eliminar caracteres restantes

    texto =
        texto.replace(
            /[^0-9.-]/g,
            ""
        );


    const numero =
        Number(texto);


    if (
        !Number.isFinite(numero)
    ) {

        return 0;

    }


    return numero;

}


// ============================================================
// OBTENER AÑO
// ============================================================

function obtenerAnio(
    columnas
) {

    const valor =
        obtenerValor(
            columnas,
            [
                "ano",
                "año",
                "year"
            ]
        );


    if (
        valor !== "" &&
        valor !== null &&
        valor !== undefined
    ) {

        const numero =
            Number(valor);

        if (
            numero >= 1900 &&
            numero <= 2100
        ) {

            return numero;

        }

    }


    const fecha =
        obtenerValor(
            columnas,
            [
                "fecha",
                "fecharemision"
            ]
        );


    const fechaConvertida =
        convertirFecha(
            fecha
        );


    if (
        fechaConvertida
    ) {

        return fechaConvertida.getFullYear();

    }


    return 2026;

}


// ============================================================
// OBTENER MES
// ============================================================

function obtenerMes(
    columnas
) {

    const valor =
        obtenerValor(
            columnas,
            [
                "mes"
            ]
        );


    if (
        valor !== "" &&
        valor !== null &&
        valor !== undefined
    ) {

        const numero =
            Number(valor);

        if (
            numero >= 1 &&
            numero <= 12
        ) {

            return numero;

        }

    }


    const fecha =
        obtenerValor(
            columnas,
            [
                "fecha",
                "fecharemision"
            ]
        );


    const fechaConvertida =
        convertirFecha(
            fecha
        );


    if (
        fechaConvertida
    ) {

        return (
            fechaConvertida.getMonth()
            + 1
        );

    }


    return 1;

}


// ============================================================
// CONVERTIR FECHA
// ============================================================

function convertirFecha(
    valor
) {

    if (
        valor instanceof Date &&
        !isNaN(valor)
    ) {

        return valor;

    }


    if (
        typeof valor === "number"
    ) {

        try {

            const fecha =
                XLSX.SSF.parse_date_code(
                    valor
                );

            if (fecha) {

                return new Date(
                    fecha.y,
                    fecha.m - 1,
                    fecha.d
                );

            }

        }

        catch (error) {

            return null;

        }

    }


    if (!valor) {
        return null;
    }


    const texto =
        String(valor)
            .trim();


    // dd/mm/yyyy

    let partes =
        texto.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
        );


    if (partes) {

        return new Date(
            Number(partes[3]),
            Number(partes[2]) - 1,
            Number(partes[1])
        );

    }


    const fecha =
        new Date(texto);


    if (!isNaN(fecha)) {

        return fecha;

    }


    return null;

}


// ============================================================
// FILTROS
// ============================================================

function activarFiltros() {

    const filtros = [
        "filtroAnio",
        "filtroMes",
        "filtroVendedor",
        "filtroCliente",
        "filtroLinea",
        "filtroProducto"
    ];


    filtros.forEach(
        function (id) {

            const elemento =
                document.getElementById(
                    id
                );

            if (elemento) {

                elemento.addEventListener(
                    "change",
                    actualizarDashboard
                );

            }

        }
    );

}


// ============================================================
// LLENAR FILTROS
// ============================================================

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


// ============================================================
// CREAR SELECT
// ============================================================

function llenarSelect(
    id,
    valores,
    textoTodos
) {

    const select =
        document.getElementById(
            id
        );


    if (!select) {
        return;
    }


    select.innerHTML = "";


    const opcionTodos =
        document.createElement(
            "option"
        );


    opcionTodos.value =
        "todos";


    opcionTodos.textContent =
        textoTodos;


    select.appendChild(
        opcionTodos
    );


    const unicos =
        [
            ...new Set(
                valores
                    .filter(
                        v =>
                            v !== "" &&
                            v !== null &&
                            v !== undefined
                    )
                    .map(
                        v =>
                            String(v)
                    )
            )
        ];


    unicos.sort(
        function (a, b) {

            return a.localeCompare(
                b,
                "es",
                {
                    numeric: true
                }
            );

        }
    );


    unicos.forEach(
        function (valor) {

            const opcion =
                document.createElement(
                    "option"
                );


            opcion.value =
                valor;


            opcion.textContent =
                valor;


            select.appendChild(
                opcion
            );

        }
    );

}


// ============================================================
// ACTUALIZAR DASHBOARD
// ============================================================

function actualizarDashboard() {

    datosFiltrados =
        datosOriginales.filter(
            function (dato) {

                return
                    coincideFiltro(
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


// ============================================================
// COINCIDENCIA FILTRO
// ============================================================

function coincideFiltro(
    valor,
    id
) {

    const filtro =
        document.getElementById(
            id
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


// ============================================================
// KPIs
// ============================================================

function actualizarKPIs() {

    let ventas = 0;
    let costos = 0;
    let utilidad = 0;


    datosFiltrados.forEach(
        function (dato) {

            ventas +=
                Number(
                    dato.valorVenta
                ) || 0;


            costos +=
                Number(
                    dato.costo
                ) || 0;


            utilidad +=
                Number(
                    dato.utilidad
                ) || 0;

        }
    );


    const margen =
        ventas !== 0
            ? utilidad / ventas
            : 0;


    const elementoVentas =
        document.getElementById(
            "ventas"
        );


    const elementoCostos =
        document.getElementById(
            "costos"
        );


    const elementoUtilidad =
        document.getElementById(
            "utilidad"
        );


    const elementoMargen =
        document.getElementById(
            "margen"
        );


    if (elementoVentas) {

        elementoVentas.textContent =
            formatoMoneda(
                ventas
            );

    }


    if (elementoCostos) {

        elementoCostos.textContent =
            formatoMoneda(
                costos
            );

    }


    if (elementoUtilidad) {

        elementoUtilidad.textContent =
            formatoMoneda(
                utilidad
            );

    }


    if (elementoMargen) {

        elementoMargen.textContent =
            formatoPorcentaje(
                margen
            );

    }

}


// ============================================================
// AGRUPAR
// ============================================================

function agrupar(
    campo
) {

    const resultado = {};


    datosFiltrados.forEach(
        function (dato) {

            const clave =
                dato[campo] ||
                "Sin información";


            if (
                !resultado[clave]
            ) {

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


// ============================================================
// GRÁFICOS
// ============================================================

function actualizarGraficos() {

    crearGraficoClientes();

    crearGraficoLineas();

    crearGraficoMensual();

    crearGraficoVendedores();

}


// ============================================================
// CLIENTES
// ============================================================

function crearGraficoClientes() {

    const canvas =
        document.getElementById(
            "graficoClientes"
        );


    if (!canvas) {
        return;
    }


    const datos =
        agrupar(
            "cliente"
        );


    const elementos =
        Object.entries(
            datos
        )
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
        )
        .slice(
            0,
            15
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

                    labels:
                        elementos.map(
                            x => x[0]
                        ),

                    datasets: [

                        {

                            label:
                                "Utilidad",

                            data:
                                elementos.map(
                                    x =>
                                        x[1].utilidad
                                )

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


// ============================================================
// LÍNEAS
// ============================================================

function crearGraficoLineas() {

    const canvas =
        document.getElementById(
            "graficoLineas"
        );


    if (!canvas) {
        return;
    }


    const datos =
        agrupar(
            "linea"
        );


    const elementos =
        Object.entries(
            datos
        )
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
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

                    labels:
                        elementos.map(
                            x => x[0]
                        ),

                    datasets: [

                        {

                            label:
                                "Utilidad",

                            data:
                                elementos.map(
                                    x =>
                                        x[1].utilidad
                                )

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


// ============================================================
// MENSUAL
// ============================================================

function crearGraficoMensual() {

    const canvas =
        document.getElementById(
            "graficoMensual"
        );


    if (!canvas) {
        return;
    }


    const meses =
        [
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
        new Array(
            12
        ).fill(0);


    datosFiltrados.forEach(
        function (dato) {

            const mes =
                Number(
                    dato.mes
                );


            if (
                mes >= 1 &&
                mes <= 12
            ) {

                valores[
                    mes - 1
                ] +=
                    Number(
                        dato.utilidad
                    ) || 0;

            }

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
                        meses,

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


// ============================================================
// VENDEDORES
// ============================================================

function crearGraficoVendedores() {

    const canvas =
        document.getElementById(
            "graficoVendedores"
        );


    if (!canvas) {
        return;
    }


    const datos =
        agrupar(
            "vendedor"
        );


    const elementos =
        Object.entries(
            datos
        )
        .sort(
            (a, b) =>
                b[1].utilidad -
                a[1].utilidad
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

                    labels:
                        elementos.map(
                            x => x[0]
                        ),

                    datasets: [

                        {

                            label:
                                "Utilidad",

                            data:
                                elementos.map(
                                    x =>
                                        x[1].utilidad
                                )

                        }

                    ]

                },

                options: {

                    indexAxis:
                        "y",

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


// ============================================================
// TABLA PRODUCTOS
// ============================================================

function actualizarTabla() {

    const cuerpo =
        document.getElementById(
            "tablaProductos"
        );


    if (!cuerpo) {
        return;
    }


    cuerpo.innerHTML = "";


    const datos =
        agrupar(
            "producto"
        );


    const productos =
        Object.entries(
            datos
        )
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
        function (
            [producto, datos]
        ) {

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
                    ${escapeHTML(
                        producto
                    )}
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


// ============================================================
// FORMATO MONEDA
// ============================================================

function formatoMoneda(
    valor
) {

    return new Intl.NumberFormat(
        "es-CO",
        {

            style:
                "currency",

            currency:
                "COP",

            maximumFractionDigits:
                0

        }
    ).format(
        Number(valor) || 0
    );

}


// ============================================================
// FORMATO PORCENTAJE
// ============================================================

function formatoPorcentaje(
    valor
) {

    return new Intl.NumberFormat(
        "es-CO",
        {

            style:
                "percent",

            minimumFractionDigits:
                1,

            maximumFractionDigits:
                1

        }
    ).format(
        Number(valor) || 0
    );

}


// ============================================================
// SEGURIDAD HTML
// ============================================================

function escapeHTML(
    valor
) {

    return String(
        valor
    )
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
