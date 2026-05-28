import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import emailjs from "@emailjs/browser";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://wjbmjauvwyseugtaugjb.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_exynDYj6j9ZtXOwhiAC7vQ_dD5Swvu2";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const entrenadores = {
  "Tercera Division": [
    "gbuenosc@gmail.com",
    "hugoguti05@gmail.com",
    "fgarciadt91@gmail.com",
    "videoanalisisracing2022@gmail.com",
  ],
  "Sub 19": [
    "videoanalisisracing2022@gmail.com",
    "gbuenosc@gmail.com",
    "sebasmurgia75@gmail.com",
  ],
  "Sub 17": [
    "ernek17@gmail.com",
    "videoanalisisracing2022@gmail.com",
    "liberhistorico11@gmail.com",
  ],
  "Sub 16": [
    "videoanalisisracing2022@gmail.com",
    "fgarciadt91@gmail.com",
    "jimymachado95@gmail.com",
  ],
  "Sub 15": [
    "videoanalisisracing2022@gmail.com",
    "guillermodavidcalcastellano@gmail.com",
    "mauriciomassa7@gmail.com",
    "maxi1524@gmail.com",
  ],
  "Sub 14": [
    "videoanalisisracing2022@gmail.com",
    "icardorulo@gmail.com",
    "martinbonino14@gmail.com",
    "maxi1524@gmail.com",
  ],
};

export default function App() {
  const [jugadores, setJugadores] = useState([]);
  const [busqueda, setBusqueda] = useState("");

  const [categoriaSeleccionada, setCategoriaSeleccionada] =
    useState("Todas");

  const [prioridadSeleccionada, setPrioridadSeleccionada] =
    useState("Todas");

  const [categoriaAlertas, setCategoriaAlertas] =
    useState("Todas");

  const [mostrarVistaPrevia, setMostrarVistaPrevia] =
    useState(false);

  const [mostrarVistaPreviaNuevas, setMostrarVistaPreviaNuevas] =
    useState(false);

  const [avisados, setAvisados] = useState([]);

  const [textoConvocados, setTextoConvocados] = useState("");
  const [resultadosConvocados, setResultadosConvocados] = useState([]);

  useEffect(() => {
    cargarExcel();
    cargarAvisados();
  }, []);

  async function cargarAvisados() {
    const { data, error } = await supabase
      .from("avisos_fichas_medicas")
      .select("clave");

    if (error) {
      console.error("Error cargando avisos:", error);
      return;
    }

    const claves = data.map((aviso) => aviso.clave);

    setAvisados(claves);
  }

  async function registrarAvisados(alertas) {
    const registros = alertas.flatMap((alerta) =>
      alerta.jugadores.map((jugador) => ({
        clave: obtenerClaveAviso(jugador),
        categoria: jugador.categoria,
        cedula: String(jugador.cedula),
        nombre: jugador.nombre,
        vencimiento: jugador.vencimiento,
      }))
    );

    if (registros.length === 0) return;

    const { error } = await supabase
      .from("avisos_fichas_medicas")
      .upsert(registros, {
        onConflict: "clave",
        ignoreDuplicates: true,
      });

    if (error) {
      console.error("Error registrando avisos:", error);
      alert("Hubo un error guardando los avisos.");
      return;
    }

    await cargarAvisados();
  }

  async function reiniciarAvisos() {
    const textoCategoria =
      categoriaAlertas === "Todas"
        ? "todas las categorías"
        : categoriaAlertas;

    const confirmar = window.confirm(
      `Esto borrará el registro online de avisos enviados para ${textoCategoria}. Las urgencias actuales volverán a aparecer como nuevas. ¿Confirmás?`
    );

    if (!confirmar) return;

    let query = supabase.from("avisos_fichas_medicas").delete();

    if (categoriaAlertas === "Todas") {
      query = query.neq("clave", "");
    } else {
      query = query.eq("categoria", categoriaAlertas);
    }

    const { error } = await query;

    if (error) {
      console.error("Error reiniciando avisos:", error);
      alert("Error al reiniciar avisos.");
      return;
    }

    await cargarAvisados();

    alert(`Registro online de avisos reiniciado para ${textoCategoria}.`);
  }

  async function marcarNuevasComoAvisadas() {
    const alertas = prepararAlertas(true);

    if (alertas.length === 0) {
      alert("No hay nuevas urgencias para marcar como avisadas.");
      return;
    }

    const cantidadJugadores = alertas.reduce(
      (total, alerta) => total + alerta.jugadores.length,
      0
    );

    const textoCategoria =
      categoriaAlertas === "Todas"
        ? "todas las categorías"
        : categoriaAlertas;

    const confirmar = window.confirm(
      `Se marcarán ${cantidadJugadores} nuevas urgencias como avisadas en ${textoCategoria}, sin enviar correos. ¿Confirmás?`
    );

    if (!confirmar) return;

    await registrarAvisados(alertas);

    alert(
      `Nuevas urgencias marcadas como avisadas: ${cantidadJugadores}`
    );
  }

  async function cargarExcel() {
    const response = await fetch(
      `/jugadores.xlsx?v=${new Date().getTime()}`
    );

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data);

    let jugadoresProcesados = [];

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const filas = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      });

      for (let i = 0; i < filas.length; i++) {
        const fila = filas[i];

        if (!fila || fila.length === 0) continue;

        if (
          String(fila[0]).toLowerCase().includes("nombre") ||
          String(fila[0]).toLowerCase().includes("tercera") ||
          String(fila[0]).toLowerCase().includes("sub")
        ) {
          continue;
        }

        if (
          typeof fila[0] === "string" &&
          fila[1] &&
          fila[2]
        ) {
          jugadoresProcesados.push({
            id: `${sheetName}-${i}`,
            nombre: fila[0],
            cedula: fila[1],
            vencimiento: formatearFecha(fechaDesdeExcel(fila[2])),
            estadoExcel: fila[3] || "",
            categoria: sheetName,
          });
        }
      }
    });

    setJugadores(jugadoresProcesados);
  }

  function fechaDesdeExcel(fechaExcel) {
    if (typeof fechaExcel === "number") {
      const fecha = XLSX.SSF.parse_date_code(fechaExcel);

      return `${fecha.y}-${String(fecha.m).padStart(
        2,
        "0"
      )}-${String(fecha.d).padStart(2, "0")}`;
    }

    return fechaExcel;
  }

  function formatearFecha(fecha) {
    if (!fecha) return "";

    if (typeof fecha === "string" && fecha.includes("/")) {
      const partes = fecha.split("/");

      if (partes.length === 3) {
        const dia = partes[0].padStart(2, "0");
        const mes = partes[1].padStart(2, "0");
        const anio = partes[2];

        return `${anio}-${mes}-${dia}`;
      }
    }

    return fecha;
  }

  function crearFechaLocal(fecha) {
    if (!fecha) return null;

    if (typeof fecha === "string" && fecha.includes("-")) {
      const [anio, mes, dia] = fecha.split("-").map(Number);

      return new Date(anio, mes - 1, dia);
    }

    if (typeof fecha === "string" && fecha.includes("/")) {
      const [dia, mes, anio] = fecha.split("/").map(Number);

      return new Date(anio, mes - 1, dia);
    }

    return new Date(fecha);
  }

  function mostrarFechaEspanol(fecha) {
    const fechaLocal = crearFechaLocal(fecha);

    if (!fechaLocal || isNaN(fechaLocal)) return fecha;

    const dia = String(fechaLocal.getDate()).padStart(2, "0");
    const mes = String(fechaLocal.getMonth() + 1).padStart(2, "0");
    const anio = fechaLocal.getFullYear();

    return `${dia}/${mes}/${anio}`;
  }

  function calcularDiasRestantes(fecha) {
    const hoy = new Date();

    const hoyLocal = new Date(
      hoy.getFullYear(),
      hoy.getMonth(),
      hoy.getDate()
    );

    const vencimiento = crearFechaLocal(fecha);

    if (!vencimiento || isNaN(vencimiento)) return 9999;

    const diferencia = vencimiento - hoyLocal;

    return Math.ceil(
      diferencia / (1000 * 60 * 60 * 24)
    );
  }

  function obtenerEstado(dias) {
    if (dias < 0) {
      return {
        texto: "Vencida",
        color: "#111111",
        prioridad: 1,
      };
    }

    if (dias <= 45) {
      return {
        texto: "Urgente",
        color: "#dc2626",
        prioridad: 2,
      };
    }

    if (dias <= 90) {
      return {
        texto: "Atención",
        color: "#f59e0b",
        prioridad: 3,
      };
    }

    return {
      texto: "Al día",
      color: "#16a34a",
      prioridad: 4,
    };
  }

  function obtenerIconoAlerta(dias) {
    if (dias < 0) return "⛔";
    return "🔴";
  }

  function obtenerTextoDias(dias) {
    if (dias < 0) {
      return `vencida hace ${Math.abs(dias)} días`;
    }

    if (dias === 0) {
      return "vence hoy";
    }

    if (dias === 1) {
      return "vence mañana";
    }

    return `vence en ${dias} días`;
  }

  function obtenerClaveAviso(jugador) {
    return `${jugador.categoria}-${jugador.cedula}-${jugador.vencimiento}`;
  }

  function jugadorYaAvisado(jugador) {
    const clave = obtenerClaveAviso(jugador);

    return avisados.includes(clave);
  }

  function obtenerCategoriasParaAlertas() {
    const categorias = Object.keys(entrenadores);

    if (categoriaAlertas === "Todas") {
      return categorias;
    }

    return categorias.filter((categoria) => categoria === categoriaAlertas);
  }

  function prepararAlertas(soloNuevas = false) {
    let alertas = [];

    const categoriasParaAlertas = obtenerCategoriasParaAlertas();

    categoriasParaAlertas.forEach((categoria) => {
      const mailsEntrenadores = entrenadores[categoria];

      const jugadoresCategoria = jugadores
        .filter((jugador) => {
          if (jugador.categoria !== categoria) {
            return false;
          }

          const dias = calcularDiasRestantes(
            jugador.vencimiento
          );

          const estado = obtenerEstado(dias);

          const esAlerta =
            estado.texto === "Vencida" ||
            estado.texto === "Urgente";

          if (!esAlerta) {
            return false;
          }

          if (soloNuevas && jugadorYaAvisado(jugador)) {
            return false;
          }

          return true;
        })
        .sort((a, b) => {
          const diasA = calcularDiasRestantes(
            a.vencimiento
          );

          const diasB = calcularDiasRestantes(
            b.vencimiento
          );

          return diasA - diasB;
        });

      if (jugadoresCategoria.length === 0) {
        return;
      }

      const mensajeJugadores = jugadoresCategoria
        .map((jugador) => {
          const dias = calcularDiasRestantes(
            jugador.vencimiento
          );

          const icono = obtenerIconoAlerta(dias);
          const textoDias = obtenerTextoDias(dias);

          return `${icono} ${jugador.nombre} - ${textoDias} (${mostrarFechaEspanol(
            jugador.vencimiento
          )})`;
        })
        .join("\n");

      const subject = soloNuevas
        ? `Nuevas urgencias - Fichas médicas - ${categoria}`
        : `Control de fichas médicas - ${categoria}`;

      const mensaje = `
Categoría: ${categoria}

${mensajeJugadores}

Por favor gestionar renovaciones correspondientes.
      `;

      alertas.push({
        categoria,
        mailsEntrenadores,
        jugadores: jugadoresCategoria,
        subject,
        mensaje,
      });
    });

    return alertas;
  }

  async function enviarAlertas(soloNuevas = false) {
    const alertas = prepararAlertas(soloNuevas);

    if (alertas.length === 0) {
      alert(
        soloNuevas
          ? "No hay nuevas urgencias para enviar en la categoría seleccionada."
          : "No hay jugadores vencidos o urgentes para enviar en la categoría seleccionada."
      );
      return;
    }

    const cantidadCorreos = alertas.reduce(
      (total, alerta) =>
        total + alerta.mailsEntrenadores.length,
      0
    );

    const textoCategoria =
      categoriaAlertas === "Todas"
        ? "todas las categorías"
        : categoriaAlertas;

    const confirmar = window.confirm(
      soloNuevas
        ? `Se enviarán ${cantidadCorreos} correos con NUEVAS urgencias en ${textoCategoria}. ¿Confirmás el envío?`
        : `Se enviarán ${cantidadCorreos} correos de alerta en ${textoCategoria}. ¿Confirmás el envío?`
    );

    if (!confirmar) return;

    try {
      for (const alerta of alertas) {
        for (const mail of alerta.mailsEntrenadores) {
          await emailjs.send(
            "service_5x2tbn1",
            "template_ud6605j",
            {
              subject: alerta.subject,
              mensaje: alerta.mensaje,
              to_email: mail,
            },
            "762iROv7nnKZVp4Hc"
          );
        }
      }

      if (soloNuevas) {
        await registrarAvisados(alertas);
      }

      alert(
        soloNuevas
          ? `Nuevas urgencias enviadas correctamente. Correos enviados: ${cantidadCorreos}`
          : `Alertas enviadas correctamente. Correos enviados: ${cantidadCorreos}`
      );
    } catch (error) {
      console.error(error);
      alert("Error al enviar alertas.");
    }
  }

  function normalizarTexto(texto) {
    return String(texto || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9ñ\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function separarLineasConvocados(texto) {
    return texto
      .split(/\n|,|;/)
      .map((linea) => linea.trim())
      .filter((linea) => linea.length > 0);
  }

  function buscarJugadorConvocado(nombreIngresado) {
    const buscado = normalizarTexto(nombreIngresado);

    if (!buscado) return null;

    const coincidenciaExacta = jugadores.find(
      (jugador) => normalizarTexto(jugador.nombre) === buscado
    );

    if (coincidenciaExacta) {
      return {
        jugador: coincidenciaExacta,
        tipo: "Coincidencia exacta",
      };
    }

    const coincidenciaParcial = jugadores.find((jugador) => {
      const nombreJugador = normalizarTexto(jugador.nombre);

      return (
        nombreJugador.includes(buscado) ||
        buscado.includes(nombreJugador)
      );
    });

    if (coincidenciaParcial) {
      return {
        jugador: coincidenciaParcial,
        tipo: "Coincidencia parcial",
      };
    }

    const palabrasBuscado = buscado
      .split(" ")
      .filter((palabra) => palabra.length >= 3);

    const coincidenciaPorPalabras = jugadores.find((jugador) => {
      const nombreJugador = normalizarTexto(jugador.nombre);

      return palabrasBuscado.every((palabra) =>
        nombreJugador.includes(palabra)
      );
    });

    if (coincidenciaPorPalabras) {
      return {
        jugador: coincidenciaPorPalabras,
        tipo: "Coincidencia por palabras",
      };
    }

    return null;
  }

  function verificarConvocados() {
    const nombres = separarLineasConvocados(textoConvocados);

    if (nombres.length === 0) {
      alert("Pegá primero una lista de convocados.");
      return;
    }

    const resultados = nombres.map((nombreOriginal) => {
      const encontrado = buscarJugadorConvocado(nombreOriginal);

      if (!encontrado) {
        return {
          nombreOriginal,
          encontrado: false,
        };
      }

      const jugador = encontrado.jugador;
      const dias = calcularDiasRestantes(jugador.vencimiento);
      const estado = obtenerEstado(dias);

      return {
        nombreOriginal,
        encontrado: true,
        jugador,
        dias,
        estado,
        tipo: encontrado.tipo,
      };
    });

    setResultadosConvocados(resultados);
  }

  function limpiarConvocados() {
    setTextoConvocados("");
    setResultadosConvocados([]);
  }

  function copiarResumenConvocados() {
    if (resultadosConvocados.length === 0) {
      alert("Primero verificá una lista de convocados.");
      return;
    }

    const noHabilitados = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Vencida"
    );

    const urgentes = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Urgente"
    );

    const noEncontrados = resultadosConvocados.filter(
      (resultado) => !resultado.encontrado
    );

    let mensaje = "Revisión de fichas médicas:\n\n";

    if (noHabilitados.length === 0) {
      mensaje += "✅ No hay convocados con ficha vencida.\n\n";
    } else {
      mensaje += "⛔ Convocados con ficha vencida:\n";
      noHabilitados.forEach((resultado) => {
        mensaje += `• ${resultado.jugador.nombre} - ${obtenerTextoDias(
          resultado.dias
        )} (${mostrarFechaEspanol(resultado.jugador.vencimiento)})\n`;
      });
      mensaje += "\n";
    }

    if (urgentes.length > 0) {
      mensaje += "🔴 Convocados con ficha urgente:\n";
      urgentes.forEach((resultado) => {
        mensaje += `• ${resultado.jugador.nombre} - ${obtenerTextoDias(
          resultado.dias
        )} (${mostrarFechaEspanol(resultado.jugador.vencimiento)})\n`;
      });
      mensaje += "\n";
    }

    if (noEncontrados.length > 0) {
      mensaje += "❓ No encontrados en el listado:\n";
      noEncontrados.forEach((resultado) => {
        mensaje += `• ${resultado.nombreOriginal}\n`;
      });
    }

    navigator.clipboard.writeText(mensaje);

    alert("Resumen copiado al portapapeles.");
  }

  const alertasPreview = useMemo(() => {
    return prepararAlertas(false);
  }, [jugadores, avisados, categoriaAlertas]);

  const nuevasUrgenciasPreview = useMemo(() => {
    return prepararAlertas(true);
  }, [jugadores, avisados, categoriaAlertas]);

  const cantidadNuevasUrgencias = useMemo(() => {
    return nuevasUrgenciasPreview.reduce(
      (total, alerta) => total + alerta.jugadores.length,
      0
    );
  }, [nuevasUrgenciasPreview]);

  const resumenConvocados = useMemo(() => {
    const vencidos = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Vencida"
    ).length;

    const urgentes = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Urgente"
    ).length;

    const atencion = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Atención"
    ).length;

    const alDia = resultadosConvocados.filter(
      (resultado) =>
        resultado.encontrado &&
        resultado.estado.texto === "Al día"
    ).length;

    const noEncontrados = resultadosConvocados.filter(
      (resultado) => !resultado.encontrado
    ).length;

    return {
      vencidos,
      urgentes,
      atencion,
      alDia,
      noEncontrados,
    };
  }, [resultadosConvocados]);

  const jugadoresFiltrados = useMemo(() => {
    return [...jugadores]
      .filter((jugador) => {
        const coincideBusqueda =
          jugador.nombre
            .toLowerCase()
            .includes(busqueda.toLowerCase());

        const coincideCategoria =
          categoriaSeleccionada === "Todas" ||
          jugador.categoria === categoriaSeleccionada;

        const dias = calcularDiasRestantes(
          jugador.vencimiento
        );

        const estado = obtenerEstado(dias);

        const coincidePrioridad =
          prioridadSeleccionada === "Todas" ||
          estado.texto === prioridadSeleccionada;

        return (
          coincideBusqueda &&
          coincideCategoria &&
          coincidePrioridad
        );
      })
      .sort((a, b) => {
        const estadoA = obtenerEstado(
          calcularDiasRestantes(a.vencimiento)
        );

        const estadoB = obtenerEstado(
          calcularDiasRestantes(b.vencimiento)
        );

        return estadoA.prioridad - estadoB.prioridad;
      });
  }, [
    jugadores,
    busqueda,
    categoriaSeleccionada,
    prioridadSeleccionada,
  ]);

  const dashboard = useMemo(() => {
    let vencidas = 0;
    let urgentes = 0;
    let atencion = 0;
    let alDia = 0;

    jugadoresFiltrados.forEach((jugador) => {
      const dias = calcularDiasRestantes(
        jugador.vencimiento
      );

      const estado = obtenerEstado(dias);

      if (estado.texto === "Vencida") {
        vencidas++;
      } else if (estado.texto === "Urgente") {
        urgentes++;
      } else if (estado.texto === "Atención") {
        atencion++;
      } else {
        alDia++;
      }
    });

    return {
      vencidas,
      urgentes,
      atencion,
      alDia,
    };
  }, [jugadoresFiltrados]);

  const textoCategoriaAlertas =
    categoriaAlertas === "Todas"
      ? "todas las categorías"
      : categoriaAlertas;

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#eef2f7",
        padding: "20px",
        fontFamily:
          "-apple-system, BlinkMacSystemFont, sans-serif",
      }}
    >
      <div
        style={{
          maxWidth: "1400px",
          margin: "0 auto",
        }}
      >
        <div
          style={{
            marginBottom: "30px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: "15px",
          }}
        >
          <div>
            <h1
              style={{
                fontSize: "36px",
                marginBottom: "10px",
                color: "#111827",
              }}
            >
              Control de Fichas Médicas
            </h1>

            <p
              style={{
                color: "#6b7280",
                fontSize: "16px",
              }}
            >
              Seguimiento de vencimientos y
              habilitaciones de jugadores
            </p>

            <p
              style={{
                color:
                  cantidadNuevasUrgencias > 0
                    ? "#dc2626"
                    : "#16a34a",
                fontWeight: "bold",
                marginTop: "8px",
              }}
            >
              🔔 Nuevas urgencias detectadas en{" "}
              {textoCategoriaAlertas}: {cantidadNuevasUrgencias}
            </p>
          </div>

          <div
            style={{
              display: "flex",
              gap: "10px",
              flexWrap: "wrap",
              alignItems: "center",
            }}
          >
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                gap: "4px",
              }}
            >
              <label
                style={{
                  fontSize: "13px",
                  color: "#374151",
                  fontWeight: "bold",
                }}
              >
                Categoría para alertas
              </label>

              <select
                value={categoriaAlertas}
                onChange={(e) =>
                  setCategoriaAlertas(e.target.value)
                }
                style={{
                  padding: "12px",
                  borderRadius: "12px",
                  border: "1px solid #d1d5db",
                  fontWeight: "bold",
                  minWidth: "190px",
                  cursor: "pointer",
                }}
              >
                <option value="Todas">Todas</option>

                {Object.keys(entrenadores).map((categoria) => (
                  <option key={categoria} value={categoria}>
                    {categoria}
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={() =>
                setMostrarVistaPrevia(
                  !mostrarVistaPrevia
                )
              }
              style={buttonDark}
            >
              👁️ Vista previa
            </button>

            <button
              onClick={() =>
                setMostrarVistaPreviaNuevas(
                  !mostrarVistaPreviaNuevas
                )
              }
              style={buttonBlue}
            >
              🔔 Nuevas urgencias
            </button>

            <button
              onClick={marcarNuevasComoAvisadas}
              style={buttonGreen}
            >
              ✅ Marcar nuevas como avisadas
            </button>

            <button
              onClick={() => enviarAlertas(true)}
              style={buttonPurple}
            >
              📩 Enviar nuevas
            </button>

            <button
              onClick={() => enviarAlertas(false)}
              style={buttonRed}
            >
              🚨 Enviar todas
            </button>

            <button
              onClick={reiniciarAvisos}
              style={buttonGray}
            >
              ♻️ Reiniciar avisos
            </button>
          </div>
        </div>

        {mostrarVistaPrevia && (
          <VistaPreviaAlertas
            titulo={`Vista previa de todas las alertas - ${textoCategoriaAlertas}`}
            alertas={alertasPreview}
            calcularDiasRestantes={calcularDiasRestantes}
            obtenerIconoAlerta={obtenerIconoAlerta}
            obtenerTextoDias={obtenerTextoDias}
            mostrarFechaEspanol={mostrarFechaEspanol}
          />
        )}

        {mostrarVistaPreviaNuevas && (
          <VistaPreviaAlertas
            titulo={`Vista previa de nuevas urgencias - ${textoCategoriaAlertas}`}
            alertas={nuevasUrgenciasPreview}
            calcularDiasRestantes={calcularDiasRestantes}
            obtenerIconoAlerta={obtenerIconoAlerta}
            obtenerTextoDias={obtenerTextoDias}
            mostrarFechaEspanol={mostrarFechaEspanol}
          />
        )}

        <div
          style={{
            display: "grid",
            gridTemplateColumns:
              "repeat(auto-fit, minmax(220px, 1fr))",
            gap: "18px",
            marginBottom: "30px",
          }}
        >
          <DashboardCard
            titulo="⚫ Vencidas"
            cantidad={dashboard.vencidas}
            color="#111111"
          />

          <DashboardCard
            titulo="🔴 Urgentes"
            cantidad={dashboard.urgentes}
            color="#dc2626"
          />

          <DashboardCard
            titulo="🟠 Atención"
            cantidad={dashboard.atencion}
            color="#f59e0b"
          />

          <DashboardCard
            titulo="🟢 Al día"
            cantidad={dashboard.alDia}
            color="#16a34a"
          />
        </div>

        <div
          style={{
            background: "white",
            borderRadius: "18px",
            padding: "20px",
            marginBottom: "30px",
            boxShadow:
              "0 4px 20px rgba(0,0,0,0.06)",
          }}
        >
          <h2
            style={{
              marginTop: 0,
              color: "#111827",
            }}
          >
            Verificación de convocados
          </h2>

          <p
            style={{
              color: "#6b7280",
              marginTop: 0,
            }}
          >
            Pegá la lista de convocados, un jugador por línea, y la app verificará su ficha médica.
          </p>

          <textarea
            value={textoConvocados}
            onChange={(e) => setTextoConvocados(e.target.value)}
            placeholder={`Ejemplo:\nJuan Pérez\nLucas Rodríguez\nMartín Silva`}
            style={{
              width: "100%",
              minHeight: "160px",
              padding: "14px",
              borderRadius: "12px",
              border: "1px solid #d1d5db",
              fontSize: "15px",
              outline: "none",
              boxSizing: "border-box",
              resize: "vertical",
              marginBottom: "15px",
            }}
          />

          <div
            style={{
              display: "flex",
              gap: "10px",
              flexWrap: "wrap",
              marginBottom:
                resultadosConvocados.length > 0 ? "20px" : 0,
            }}
          >
            <button
              onClick={verificarConvocados}
              style={buttonBlue}
            >
              🔎 Verificar convocados
            </button>

            <button
              onClick={copiarResumenConvocados}
              style={buttonGreen}
            >
              📋 Copiar resumen
            </button>

            <button
              onClick={limpiarConvocados}
              style={buttonGray}
            >
              Limpiar
            </button>
          </div>

          {resultadosConvocados.length > 0 && (
            <div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns:
                    "repeat(auto-fit, minmax(160px, 1fr))",
                  gap: "12px",
                  marginBottom: "20px",
                }}
              >
                <MiniCard
                  titulo="⛔ Vencidos"
                  cantidad={resumenConvocados.vencidos}
                  color="#111111"
                />

                <MiniCard
                  titulo="🔴 Urgentes"
                  cantidad={resumenConvocados.urgentes}
                  color="#dc2626"
                />

                <MiniCard
                  titulo="🟠 Atención"
                  cantidad={resumenConvocados.atencion}
                  color="#f59e0b"
                />

                <MiniCard
                  titulo="🟢 Al día"
                  cantidad={resumenConvocados.alDia}
                  color="#16a34a"
                />

                <MiniCard
                  titulo="❓ No encontrados"
                  cantidad={resumenConvocados.noEncontrados}
                  color="#6b7280"
                />
              </div>

              <div
                style={{
                  overflowX: "auto",
                  borderRadius: "12px",
                }}
              >
                <table
                  style={{
                    width: "100%",
                    borderCollapse: "collapse",
                    minWidth: "900px",
                  }}
                >
                  <thead>
                    <tr>
                      <th style={thStyle}>Ingresado</th>
                      <th style={thStyle}>Jugador encontrado</th>
                      <th style={thStyle}>Categoría</th>
                      <th style={thStyle}>Vencimiento</th>
                      <th style={thStyle}>Días</th>
                      <th style={thStyle}>Estado</th>
                      <th style={thStyle}>Coincidencia</th>
                    </tr>
                  </thead>

                  <tbody>
                    {resultadosConvocados.map((resultado, index) => {
                      if (!resultado.encontrado) {
                        return (
                          <tr
                            key={`${resultado.nombreOriginal}-${index}`}
                            style={{
                              borderBottom: "1px solid #e5e7eb",
                            }}
                          >
                            <td style={tdStyle}>
                              <strong>{resultado.nombreOriginal}</strong>
                            </td>

                            <td style={tdStyle}>No encontrado</td>
                            <td style={tdStyle}>—</td>
                            <td style={tdStyle}>—</td>
                            <td style={tdStyle}>—</td>

                            <td style={tdStyle}>
                              <span
                                style={{
                                  background: "#6b7280",
                                  color: "white",
                                  padding: "8px 12px",
                                  borderRadius: "999px",
                                  fontWeight: "bold",
                                  fontSize: "14px",
                                  display: "inline-block",
                                  minWidth: "110px",
                                  textAlign: "center",
                                }}
                              >
                                No encontrado
                              </span>
                            </td>

                            <td style={tdStyle}>—</td>
                          </tr>
                        );
                      }

                      return (
                        <tr
                          key={`${resultado.nombreOriginal}-${index}`}
                          style={{
                            borderBottom: "1px solid #e5e7eb",
                          }}
                        >
                          <td style={tdStyle}>
                            {resultado.nombreOriginal}
                          </td>

                          <td style={tdStyle}>
                            <strong>{resultado.jugador.nombre}</strong>
                          </td>

                          <td style={tdStyle}>
                            {resultado.jugador.categoria}
                          </td>

                          <td style={tdStyle}>
                            {mostrarFechaEspanol(
                              resultado.jugador.vencimiento
                            )}
                          </td>

                          <td style={tdStyle}>{resultado.dias}</td>

                          <td style={tdStyle}>
                            <span
                              style={{
                                background: resultado.estado.color,
                                color: "white",
                                padding: "8px 12px",
                                borderRadius: "999px",
                                fontWeight: "bold",
                                fontSize: "14px",
                                display: "inline-block",
                                minWidth: "90px",
                                textAlign: "center",
                              }}
                            >
                              {resultado.estado.texto}
                            </span>
                          </td>

                          <td style={tdStyle}>{resultado.tipo}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        <div
          style={{
            background: "white",
            borderRadius: "18px",
            padding: "20px",
            boxShadow:
              "0 4px 20px rgba(0,0,0,0.06)",
          }}
        >
          <SectionTitle title="Prioridad" />

          <div style={filtrosContainer}>
            {[
              "Todas",
              "Vencida",
              "Urgente",
              "Atención",
              "Al día",
            ].map((prioridad) => (
              <FiltroButton
                key={prioridad}
                activo={
                  prioridadSeleccionada === prioridad
                }
                onClick={() =>
                  setPrioridadSeleccionada(prioridad)
                }
                texto={prioridad}
              />
            ))}
          </div>

          <SectionTitle title="Categorías" />

          <div style={filtrosContainer}>
            {[
              "Todas",
              ...new Set(
                jugadores.map((j) => j.categoria)
              ),
            ].map((categoria) => (
              <FiltroButton
                key={categoria}
                activo={
                  categoriaSeleccionada === categoria
                }
                onClick={() =>
                  setCategoriaSeleccionada(categoria)
                }
                texto={categoria}
              />
            ))}
          </div>

          <input
            type="text"
            placeholder="Buscar jugador..."
            value={busqueda}
            onChange={(e) =>
              setBusqueda(e.target.value)
            }
            style={{
              width: "100%",
              padding: "14px",
              marginBottom: "25px",
              borderRadius: "12px",
              border: "1px solid #d1d5db",
              fontSize: "16px",
              outline: "none",
              boxSizing: "border-box",
            }}
          />

          <div
            style={{
              overflowX: "auto",
              borderRadius: "12px",
            }}
          >
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                minWidth: "950px",
              }}
            >
              <thead>
                <tr>
                  <th style={thStyle}>Jugador</th>
                  <th style={thStyle}>Categoría</th>
                  <th style={thStyle}>Cédula</th>
                  <th style={thStyle}>Vencimiento</th>
                  <th style={thStyle}>Días</th>
                  <th style={thStyle}>Estado</th>
                  <th style={thStyle}>Avisado</th>
                </tr>
              </thead>

              <tbody>
                {jugadoresFiltrados.map((jugador) => {
                  const dias =
                    calcularDiasRestantes(
                      jugador.vencimiento
                    );

                  const estado = obtenerEstado(dias);
                  const yaAvisado = jugadorYaAvisado(jugador);

                  return (
                    <tr
                      key={jugador.id}
                      style={{
                        borderBottom:
                          "1px solid #e5e7eb",
                      }}
                    >
                      <td style={tdStyle}>
                        <strong>{jugador.nombre}</strong>
                      </td>

                      <td style={tdStyle}>
                        {jugador.categoria}
                      </td>

                      <td style={tdStyle}>
                        {jugador.cedula}
                      </td>

                      <td style={tdStyle}>
                        {mostrarFechaEspanol(
                          jugador.vencimiento
                        )}
                      </td>

                      <td style={tdStyle}>{dias}</td>

                      <td style={tdStyle}>
                        <span
                          style={{
                            background: estado.color,
                            color: "white",
                            padding: "8px 12px",
                            borderRadius: "999px",
                            fontWeight: "bold",
                            fontSize: "14px",
                            display: "inline-block",
                            minWidth: "90px",
                            textAlign: "center",
                          }}
                        >
                          {estado.texto}
                        </span>
                      </td>

                      <td style={tdStyle}>
                        {yaAvisado ? "✅ Sí" : "—"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function VistaPreviaAlertas({
  titulo,
  alertas,
  calcularDiasRestantes,
  obtenerIconoAlerta,
  obtenerTextoDias,
  mostrarFechaEspanol,
}) {
  return (
    <div
      style={{
        background: "white",
        borderRadius: "18px",
        padding: "20px",
        marginBottom: "30px",
        boxShadow:
          "0 4px 20px rgba(0,0,0,0.06)",
      }}
    >
      <h2
        style={{
          marginTop: 0,
          color: "#111827",
        }}
      >
        {titulo}
      </h2>

      {alertas.length === 0 ? (
        <p
          style={{
            color: "#6b7280",
            fontSize: "16px",
          }}
        >
          No hay jugadores vencidos o urgentes para enviar.
        </p>
      ) : (
        <div
          style={{
            display: "grid",
            gap: "18px",
          }}
        >
          {alertas.map((alerta) => (
            <div
              key={alerta.categoria}
              style={{
                border: "1px solid #e5e7eb",
                borderRadius: "14px",
                padding: "16px",
                background: "#f9fafb",
              }}
            >
              <h3
                style={{
                  marginTop: 0,
                  marginBottom: "8px",
                  color: "#111827",
                }}
              >
                {alerta.categoria}
              </h3>

              <p
                style={{
                  marginTop: 0,
                  color: "#6b7280",
                  fontSize: "14px",
                }}
              >
                Enviar a:{" "}
                <strong>
                  {alerta.mailsEntrenadores.join(", ")}
                </strong>
              </p>

              <p
                style={{
                  marginTop: 0,
                  color: "#374151",
                  fontSize: "14px",
                }}
              >
                Asunto: <strong>{alerta.subject}</strong>
              </p>

              <ul
                style={{
                  marginBottom: 0,
                  paddingLeft: "20px",
                }}
              >
                {alerta.jugadores.map((jugador) => {
                  const dias =
                    calcularDiasRestantes(
                      jugador.vencimiento
                    );

                  return (
                    <li
                      key={jugador.id}
                      style={{
                        marginBottom: "6px",
                      }}
                    >
                      {obtenerIconoAlerta(dias)}{" "}
                      <strong>{jugador.nombre}</strong> —{" "}
                      {obtenerTextoDias(dias)} (
                      {mostrarFechaEspanol(
                        jugador.vencimiento
                      )}
                      )
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function DashboardCard({
  titulo,
  cantidad,
  color,
}) {
  return (
    <div
      style={{
        background: color,
        color: "white",
        padding: "24px",
        borderRadius: "18px",
        boxShadow:
          "0 6px 20px rgba(0,0,0,0.15)",
      }}
    >
      <div
        style={{
          fontSize: "16px",
          opacity: 0.9,
          marginBottom: "10px",
        }}
      >
        {titulo}
      </div>

      <div
        style={{
          fontSize: "42px",
          fontWeight: "bold",
        }}
      >
        {cantidad}
      </div>
    </div>
  );
}

function MiniCard({
  titulo,
  cantidad,
  color,
}) {
  return (
    <div
      style={{
        background: color,
        color: "white",
        padding: "16px",
        borderRadius: "14px",
        fontWeight: "bold",
      }}
    >
      <div
        style={{
          fontSize: "14px",
          marginBottom: "6px",
        }}
      >
        {titulo}
      </div>

      <div
        style={{
          fontSize: "28px",
        }}
      >
        {cantidad}
      </div>
    </div>
  );
}

function SectionTitle({ title }) {
  return (
    <h3
      style={{
        marginBottom: "12px",
        color: "#374151",
      }}
    >
      {title}
    </h3>
  );
}

function FiltroButton({
  activo,
  onClick,
  texto,
}) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: "10px 16px",
        border: "none",
        borderRadius: "10px",
        cursor: "pointer",
        background: activo
          ? "#111827"
          : "#e5e7eb",
        color: activo ? "white" : "#111827",
        fontWeight: "bold",
        transition: "0.2s",
      }}
    >
      {texto}
    </button>
  );
}

const baseButton = {
  color: "white",
  border: "none",
  padding: "14px 20px",
  borderRadius: "12px",
  fontWeight: "bold",
  cursor: "pointer",
  fontSize: "15px",
};

const buttonDark = {
  ...baseButton,
  background: "#111827",
};

const buttonBlue = {
  ...baseButton,
  background: "#2563eb",
};

const buttonGreen = {
  ...baseButton,
  background: "#16a34a",
};

const buttonPurple = {
  ...baseButton,
  background: "#7c3aed",
};

const buttonRed = {
  ...baseButton,
  background: "#dc2626",
};

const buttonGray = {
  ...baseButton,
  background: "#6b7280",
};

const filtrosContainer = {
  display: "flex",
  gap: "10px",
  marginBottom: "25px",
  flexWrap: "wrap",
};

const thStyle = {
  textAlign: "left",
  padding: "16px",
  background: "#111827",
  color: "white",
  fontSize: "14px",
};

const tdStyle = {
  padding: "16px",
  fontSize: "15px",
};