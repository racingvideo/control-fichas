import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import emailjs from "@emailjs/browser";

const entrenadores = {
  "Tercera Division": "ramirodoes@gmail.com",
  "Sub 19": "ramirodoes@gmail.com",
  "Sub 17": "ramirodoes@gmail.com",
  "Sub 16": "videoanalisisracing2022@gmail.com",
  "Sub 15": "videoanalisisracing2022@gmail.com",
  "Sub 14": "ramirodoes@gmail.com",
};

export default function App() {
  const [jugadores, setJugadores] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [categoriaSeleccionada, setCategoriaSeleccionada] =
    useState("Todas");
  const [prioridadSeleccionada, setPrioridadSeleccionada] =
    useState("Todas");

  useEffect(() => {
    cargarExcel();
  }, []);

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
            vencimiento: formatearFecha(fila[2]),
            estadoExcel: fila[3] || "",
            categoria: sheetName,
          });
        }
      }
    });

    setJugadores(jugadoresProcesados);
  }

  function formatearFecha(fechaExcel) {
    if (typeof fechaExcel === "number") {
      const fecha = XLSX.SSF.parse_date_code(fechaExcel);

      return `${fecha.y}-${String(fecha.m).padStart(
        2,
        "0"
      )}-${String(fecha.d).padStart(2, "0")}`;
    }

    return fechaExcel;
  }

  function calcularDiasRestantes(fecha) {
    const hoy = new Date();
    const vencimiento = new Date(fecha);
    const diferencia = vencimiento - hoy;

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

  async function enviarAlertas() {
    try {
      let cantidadMailsEnviados = 0;

      for (const categoria in entrenadores) {
        const mailEntrenador = entrenadores[categoria];

        const jugadoresCategoria = jugadores
          .filter((jugador) => {
            if (jugador.categoria !== categoria) {
              return false;
            }

            const dias = calcularDiasRestantes(
              jugador.vencimiento
            );

            const estado = obtenerEstado(dias);

            return (
              estado.texto === "Vencida" ||
              estado.texto === "Urgente"
            );
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
          continue;
        }

        const mensaje = jugadoresCategoria
          .map((jugador) => {
            const dias = calcularDiasRestantes(
              jugador.vencimiento
            );

            const icono = obtenerIconoAlerta(dias);
            const textoDias = obtenerTextoDias(dias);

            return `${icono} ${jugador.nombre} - ${textoDias} (${jugador.vencimiento})`;
          })
          .join("\n");

        await emailjs.send(
          "service_5x2tbn1",
          "template_ud6605j",
          {
            mensaje: `
Categoría: ${categoria}

${mensaje}

Por favor gestionar renovaciones correspondientes.
            `,
            to_email: mailEntrenador,
          },
          "762iROv7nnKZVp4Hc"
        );

        cantidadMailsEnviados++;
      }

      if (cantidadMailsEnviados === 0) {
        alert(
          "No hay jugadores vencidos o urgentes para enviar."
        );
        return;
      }

      alert(
        `Alertas enviadas correctamente. Correos enviados: ${cantidadMailsEnviados}`
      );
    } catch (error) {
      console.error(error);
      alert("Error al enviar alertas.");
    }
  }

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
          </div>

          <button
            onClick={enviarAlertas}
            style={{
              background: "#dc2626",
              color: "white",
              border: "none",
              padding: "14px 20px",
              borderRadius: "12px",
              fontWeight: "bold",
              cursor: "pointer",
              fontSize: "15px",
            }}
          >
            🚨 Enviar alertas
          </button>
        </div>

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
                minWidth: "850px",
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
                </tr>
              </thead>

              <tbody>
                {jugadoresFiltrados.map((jugador) => {
                  const dias =
                    calcularDiasRestantes(
                      jugador.vencimiento
                    );

                  const estado = obtenerEstado(dias);

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
                        {jugador.vencimiento}
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