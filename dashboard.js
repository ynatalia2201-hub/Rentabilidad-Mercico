let datos = [];

let graficoClientes = null;
let graficoLineas = null;
let graficoMensual = null;
let graficoVendedores = null;


// =====================================================
// CARGAR ARCHIVO EXCEL
// =====================================================

document
    .getElementById("archivoExcel")
    .addEventListener("change", cargarExcel);


function cargarExcel(event) {

    const archivo = event.target.files[0];

    if (!archivo) {
        return;
    }

    const estado = document.getElementById("estadoArchivo");

    estado.textContent = "Leyendo archivo...";

    const lector = new FileReader();


    lector.onload = function(e) {

        try {

            const contenido = new Uint8Array(e.target.result);

            const libro = XLSX.read(contenido, {
                type: "array"
            });


            // Tomar la primera hoja del Excel

            const nombreHoja = libro.SheetNames[0];

            const hoja = libro.Sheets[nombreHoja];


            // Convertir Excel a objetos

            datos = XLSX.utils.sheet_to_json(hoja, {
                defval: ""
            });


            if (datos.length === 0) {

                estado.textContent =
                    "El archivo no contiene información.";

                return;
            }


            estado.textContent =
                `Archivo cargado correctamente · ${datos.length.toLocaleString()} registros`;


            prepararDatos();

            llenarFiltros();

            actualizarDashboard();

        }

        catch (error) {

            console.error(error);

            estado.textContent =
                "No fue posible leer el archivo Excel.";

        }

    };


    lector.readAsArrayBuffer(archivo);

}


// =====================================================
// LIMPIAR DATOS
// =====================================================

function prepararDatos() {

    datos = datos.map(fila => {

        const ventas =
            fila["VrRem"] ??
            fila["Vr Rem"] ??
            fila["VRREM"] ??
            fila["Valor Remisión"] ??
            fila["Valor Remision"] ??
            0;

        const costo =
            fila["Costo"] ??
            fila["COSTO"] ??
            fila["Costos"] ??
            0;

        return {

            ...fila,

            Vendedor:
                limpiarTexto(fila["Vendedor"]),

            Cliente:
                limpiarTexto(fila["Cliente"]),

            Linea:
                limpiarTexto(fila["Linea"]),

            Producto:
                limpiarTexto(fila["Producto"]),

            Año:
                limpiarNumero(fila["Año"]),

            Mes:
                limpiarNumero(fila["Mes"]),

            Cant:
                limpiarNumero(fila["Cant"]),

            VrRem:
                limpiarNumero(ventas),

            Costo:
                limpiarNumero(costo)

        };

    });

}

function limpiarNumero(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return 0;
    }

    if (typeof valor === "number") {
        return valor;
    }

    let texto = String(valor)
        .trim()
        .replace(/\$/g, "")
        .replace(/\s/g, "");

    /*
       Maneja formatos como:

       9,460,000
       9.460.000
       $9.460.000
       $ 9,460,000
       9460000
    */

    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {

        // Formato tipo 9.460.000,50
        texto = texto
            .replace(/\./g, "")
            .replace(",", ".");

    } else if (
        texto.includes(".")
    ) {

        const partes = texto.split(".");

        if (
            partes.length > 1 &&
            partes[partes.length - 1].length === 3
        ) {

            // 9.460.000
            texto = texto.replace(/\./g, "");

        }

    } else if (
        texto.includes(",")
    ) {

        const partes = texto.split(",");

        if (
            partes.length > 1 &&
            partes[partes.length - 1].length === 3
        ) {

            // 9,460,000
            texto = texto.replace(/,/g, "");

        } else {

            // 9460000,50
            texto = texto.replace(",", ".");

        }

    }

    const numero = Number(texto);

    return isNaN(numero) ? 0 : numero;

}


// =====================================================
// LIMPIAR NÚMEROS
// =====================================================

function limpiarNumero(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return 0;

    }


    if (typeof valor === "number") {

        return valor;

    }


    let texto = String(valor)
        .trim()
        .replace(/\$/g, "")
        .replace(/\s/g, "")
        .replace(/,/g, "");


    let numero = parseFloat(texto);


    if (isNaN(numero)) {

        return 0;

    }


    return numero;

}


// =====================================================
// OBTENER DATOS FILTRADOS
// =====================================================

function obtenerDatosFiltrados() {

    const anio =
        document.getElementById("filtroAnio").value;

    const mes =
        document.getElementById("filtroMes").value;

    const vendedor =
        document.getElementById("filtroVendedor").value;

    const cliente =
        document.getElementById("filtroCliente").value;

    const linea =
        document.getElementById("filtroLinea").value;

    const producto =
        document.getElementById("filtroProducto").value;


    return datos.filter(fila => {

        return (

            (anio === "todos" ||
                String(fila.Año) === anio)

            &&

            (mes === "todos" ||
                String(fila.Mes) === mes)

            &&

            (vendedor === "todos" ||
                fila.Vendedor === vendedor)

            &&

            (cliente === "todos" ||
                fila.Cliente === cliente)

            &&

            (linea === "todos" ||
                fila.Linea === linea)

            &&

            (producto === "todos" ||
                fila.Producto === producto)

        );

    });

}


// =====================================================
// LLENAR FILTROS
// =====================================================

function llenarFiltros() {

    llenarSelect(
        "filtroAnio",
        datos.map(x => x.Año),
        "Todos"
    );


    llenarSelect(
        "filtroMes",
        datos.map(x => x.Mes),
        "Todos"
    );


    llenarSelect(
        "filtroVendedor",
        datos.map(x => x.Vendedor),
        "Todos"
    );


    llenarSelect(
        "filtroCliente",
        datos.map(x => x.Cliente),
        "Todos"
    );


    llenarSelect(
        "filtroLinea",
        datos.map(x => x.Linea),
        "Todas"
    );


    llenarSelect(
        "filtroProducto",
        datos.map(x => x.Producto),
        "Todos"
    );

}


// =====================================================
// CREAR OPCIONES DE SELECT
// =====================================================

function llenarSelect(id, valores, textoInicial) {

    const select =
        document.getElementById(id);


    select.innerHTML = "";


    const opcionInicial =
        document.createElement("option");

    opcionInicial.value = "todos";

    opcionInicial.textContent =
        textoInicial;

    select.appendChild(opcionInicial);


    const valoresUnicos =
        [...new Set(

            valores
                .filter(x => x !== "")
                .map(x => String(x))

        )]
        .sort();


    valoresUnicos.forEach(valor => {

        const opcion =
            document.createElement("option");

        opcion.value = valor;

        opcion.textContent = valor;

        select.appendChild(opcion);

    });

}


// =====================================================
// EVENTOS DE FILTROS
// =====================================================

[
    "filtroAnio",
    "filtroMes",
    "filtroVendedor",
    "filtroCliente",
    "filtroLinea",
    "filtroProducto"

].forEach(id => {

    document
        .getElementById(id)
        .addEventListener(
            "change",
            actualizarDashboard
        );

});


// =====================================================
// ACTUALIZAR DASHBOARD
// =====================================================

function actualizarDashboard() {

    const filtrados =
        obtenerDatosFiltrados();


    actualizarKPIs(filtrados);

    actualizarGraficoClientes(filtrados);

    actualizarGraficoLineas(filtrados);

    actualizarGraficoMensual(filtrados);

    actualizarGraficoVendedores(filtrados);

    actualizarTablaProductos(filtrados);

}


// =====================================================
// KPIs
// =====================================================

function actualizarKPIs(filas) {

    let ventas = 0;

    let costos = 0;


    filas.forEach(fila => {

        ventas += fila.VrRem;

        costos += fila.Costo;

    });


    const utilidad =
        ventas - costos;


    const margen =
        ventas !== 0
            ? utilidad / ventas
            : 0;


    document.getElementById("ventas")
        .textContent = formatoMoneda(ventas);


    document.getElementById("costos")
        .textContent = formatoMoneda(costos);


    document.getElementById("utilidad")
        .textContent = formatoMoneda(utilidad);


    document.getElementById("margen")
        .textContent =
        formatoPorcentaje(margen);

}


// =====================================================
// AGRUPAR INFORMACIÓN
// =====================================================

function agrupar(filas, campo) {

    const resultado = {};


    filas.forEach(fila => {

        const clave =
            fila[campo] || "Sin información";


        if (!resultado[clave]) {

            resultado[clave] = {

                ventas: 0,

                costos: 0,

                utilidad: 0

            };

        }


        resultado[clave].ventas +=
            fila.VrRem;


        resultado[clave].costos +=
            fila.Costo;


        resultado[clave].utilidad +=
            fila.VrRem - fila.Costo;

    });


    return resultado;

}


// =====================================================
// GRAFICO CLIENTES
// =====================================================

function actualizarGraficoClientes(filas) {

    const agrupado =
        agrupar(filas, "Cliente");


    const datosOrdenados =
        Object.entries(agrupado)

            .map(([nombre, valores]) => ({

                nombre,

                utilidad:
                    valores.utilidad

            }))

            .sort(
                (a, b) =>
                    b.utilidad - a.utilidad
            )

            .slice(0, 10);


    const etiquetas =
        datosOrdenados.map(x => x.nombre);


    const valores =
        datosOrdenados.map(x => x.utilidad);


    if (graficoClientes) {

        graficoClientes.destroy();

    }


    graficoClientes =
        new Chart(

            document
                .getElementById("graficoClientes"),

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


// =====================================================
// GRAFICO LINEAS
// =====================================================

function actualizarGraficoLineas(filas) {

    const agrupado =
        agrupar(filas, "Linea");


    const etiquetas =
        Object.keys(agrupado);


    const valores =
        etiquetas.map(
            x =>
                agrupado[x].utilidad
        );


    if (graficoLineas) {

        graficoLineas.destroy();

    }


    graficoLineas =
        new Chart(

            document
                .getElementById("graficoLineas"),

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

                    responsive: true

                }

            }

        );

}


// =====================================================
// GRAFICO MENSUAL
// =====================================================

function actualizarGraficoMensual(filas) {

    const agrupado = {};


    filas.forEach(fila => {

        const mes =
            fila.Mes || "Sin mes";


        if (!agrupado[mes]) {

            agrupado[mes] = {

                ventas: 0,

                utilidad: 0

            };

        }


        agrupado[mes].ventas +=
            fila.VrRem;


        agrupado[mes].utilidad +=
            fila.VrRem - fila.Costo;

    });


    const meses =
        Object.keys(agrupado)
            .sort(
                (a, b) =>
                    Number(a) - Number(b)
            );


    const ventas =
        meses.map(
            mes =>
                agrupado[mes].ventas
        );


    const utilidad =
        meses.map(
            mes =>
                agrupado[mes].utilidad
        );


    if (graficoMensual) {

        graficoMensual.destroy();

    }


    graficoMensual =
        new Chart(

            document
                .getElementById("graficoMensual"),

            {

                type: "line",

                data: {

                    labels: meses.map(
                        mes =>
                            "Mes " + mes
                    ),

                    datasets: [

                        {

                            label:
                                "Ventas",

                            data:
                                ventas

                        },

                        {

                            label:
                                "Utilidad",

                            data:
                                utilidad

                        }

                    ]

                },

                options: {

                    responsive: true

                }

            }

        );

}


// =====================================================
// GRAFICO VENDEDORES
// =====================================================

function actualizarGraficoVendedores(filas) {

    const agrupado =
        agrupar(filas, "Vendedor");


    const datosOrdenados =
        Object.entries(agrupado)

            .map(([nombre, valores]) => ({

                nombre,

                utilidad:
                    valores.utilidad

            }))

            .sort(
                (a, b) =>
                    b.utilidad - a.utilidad
            )

            .slice(0, 10);


    const etiquetas =
        datosOrdenados.map(x => x.nombre);


    const valores =
        datosOrdenados.map(x => x.utilidad);


    if (graficoVendedores) {

        graficoVendedores.destroy();

    }


    graficoVendedores =
        new Chart(

            document
                .getElementById("graficoVendedores"),

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

                    responsive: true

                }

            }

        );

}


// =====================================================
// TABLA PRODUCTOS
// =====================================================

function actualizarTablaProductos(filas) {

    const agrupado =
        agrupar(filas, "Producto");


    const productos =
        Object.entries(agrupado)

            .map(([nombre, valores]) => {

                const margen =
                    valores.ventas !== 0
                        ? valores.utilidad /
                          valores.ventas
                        : 0;


                return {

                    nombre,

                    ventas:
                        valores.ventas,

                    costos:
                        valores.costos,

                    utilidad:
                        valores.utilidad,

                    margen

                };

            })

            .sort(
                (a, b) =>
                    b.utilidad - a.utilidad
            );


    const tbody =
        document.getElementById(
            "tablaProductos"
        );


    tbody.innerHTML = "";


    productos.forEach(producto => {

        const fila =
            document.createElement("tr");


        fila.innerHTML = `

            <td>${producto.nombre}</td>

            <td>${formatoMoneda(producto.ventas)}</td>

            <td>${formatoMoneda(producto.costos)}</td>

            <td>${formatoMoneda(producto.utilidad)}</td>

            <td>${formatoPorcentaje(producto.margen)}</td>

        `;


        tbody.appendChild(fila);

    });


    if (productos.length === 0) {

        tbody.innerHTML = `

            <tr>

                <td colspan="5">

                    No hay información
                    para los filtros seleccionados.

                </td>

            </tr>

        `;

    }

}


// =====================================================
// FORMATO MONEDA
// =====================================================

function formatoMoneda(valor) {

    return new Intl.NumberFormat(
        "es-CO",
        {

            style: "currency",

            currency: "COP",

            maximumFractionDigits: 0

        }

    ).format(valor);

}


// =====================================================
// FORMATO PORCENTAJE
// =====================================================

function formatoPorcentaje(valor) {

    return (

        (valor * 100)
            .toFixed(1)

        + "%"

    );

}
