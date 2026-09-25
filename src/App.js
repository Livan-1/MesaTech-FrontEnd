import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest, apiRequest } from './authConfig';
import { useEffect, useState } from 'react';
import Axios from 'axios';

function App() {
    const { instance, accounts } = useMsal();
    const [usuarioBackend, setUsuarioBackend] = useState(null);
    const [errorBackend, setErrorBackend] = useState(null);

    // 1. Obtenemos los datos del usuario logueado
    const correo = accounts[0]?.username;
    const rolesToken = accounts[0]?.idTokenClaims?.roles;

    // 2. Determinamos el rol leyendo directamente los claims del token de Entra ID
    const rol = rolesToken?.includes("ROLE_ADMINISTRADOR") ? "Administrador"
              : rolesToken?.includes("ROLE_OPERADOR") ? "Operador"
              : rolesToken?.includes("ROLE_CLIENTE") ? "Cliente"
              : "Desconocido";

    const iniciarSesion = () => {
        instance.loginRedirect(loginRequest)
            .catch(error => {
                console.error(error);
            });
    };

    const cerrarSesion = () => {
        instance.logoutRedirect();
    };

    useEffect(() => {
        if (accounts.length === 0) {
            return;
        }

        const obtenerUsuarioBackend = async () => {
            try {
                // Solicitar a Entra ID un access token
                const tokenResponse = await instance.acquireTokenSilent({ ...apiRequest, account: accounts[0] });
                const accessToken = tokenResponse.accessToken;

                console.log("Access Token:", accessToken);

                // Consumir servicio con el access token
                Axios.get("http://localhost:8080/api/usuario", { headers: { Authorization: `Bearer ${accessToken}` } })
                    .then((response) => {
                        console.log(response.data);
                        setUsuarioBackend(response.data);
                    })
                    .catch((error) => {
                        console.log(error);
                        setErrorBackend("Error consultando api");
                    });
            } catch (error) {
                console.log("Error obteniendo datos", error);
                setErrorBackend("No fue posible obtener el access token");
            }
        };

        obtenerUsuarioBackend();
    }, [accounts, instance]);

    return (
        <div className="container" style={{ padding: "30px" }}>
            <h1>Bienvenidos a MesaTech Cloud</h1>
            
            <UnauthenticatedTemplate>
                <p className="alert alert-danger mt-3">
                    El usuario no está autenticado.
                </p>
                <button onClick={iniciarSesion} className="btn btn-primary">
                    Iniciar sesión
                </button>
            </UnauthenticatedTemplate>

            <AuthenticatedTemplate>
                <h2 className="mt-3">Usuario autenticado con éxito</h2>
                {accounts.length > 0 && (
                    <>
                        <div className="mt-3">
                            <h4>
                                Nombre de usuario:{" "}
                                {accounts[0].name} ({correo})
                            </h4>
                            <h3>
                                Rol:{" "}
                                <span className="text-primary">{rol}</span>
                            </h3>
                        </div>

                        {/* VISTA ADMINISTRADOR */}
                        {rol === "Administrador" && (
                            <>
                                <div className="card mt-3 shadow-sm border-0">
                                    <div className="card-body">
                                        <h5 className="text-success mb-3">Solicitudes Globales</h5>
                                        <div className="table-responsive">
                                            <table className="table table-hover align-middle">
                                                <thead>
                                                    <tr>
                                                        <th>ID</th>
                                                        <th>Título</th>
                                                        <th>Estado Actual</th>
                                                        <th>Acción</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {/* Aquí se conectarán los datos del backend */}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </div>

                                <div className="card mt-4 shadow-sm border-0">
                                    <div className="card-body">
                                        <h5 className="text-success mb-3">Catálogo de Categorías y Prioridades</h5>
                                        <div className="d-flex gap-2 mb-4">
                                            <input type="text" className="form-control w-25" placeholder="Nuevo ítem..." />
                                            <select className="form-select w-25">
                                                <option value="CATEGORIA">Categoría</option>
                                                <option value="PRIORIDAD">Prioridad</option>
                                            </select>
                                            <button
                                                onClick={() => console.log("Agregar al catálogo")}
                                                className="btn btn-success">
                                                Agregar
                                            </button>
                                        </div>
                                        <div className="table-responsive w-50">
                                            <table className="table table-hover align-middle">
                                                <thead>
                                                    <tr>
                                                        <th>ID</th>
                                                        <th>Nombre</th>
                                                        <th>Tipo</th>
                                                        <th>Acción</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {/* Aquí se conectarán los datos del backend */}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </div>
                            </>
                        )}

                        {/* VISTA CLIENTE */}
                        {rol === "Cliente" && (
                            <div className="card mt-3 shadow-sm">
                                <div className="card-header bg-primary text-white">
                                    <h5 className="card-title mb-0">Crear Nueva Solicitud</h5>
                                </div>
                                <div className="card-body">
                                    <form>
                                        <div className="mb-3">
                                            <label className="form-label">Título</label>
                                            <input type="text" className="form-control" />
                                        </div>
                                        <div className="mb-3">
                                            <label className="form-label">Descripción</label>
                                            <textarea className="form-control" rows="2"></textarea>
                                        </div>
                                        <div className="row mb-3">
                                            <div className="col">
                                                <label className="form-label">Categoría</label>
                                                <select className="form-select">
                                                    <option value="">Seleccione...</option>
                                                    <option value="1">Hardware</option>
                                                    <option value="2">Software</option>
                                                </select>
                                            </div>
                                            <div className="col">
                                                <label className="form-label">Prioridad</label>
                                                <select className="form-select">
                                                    <option value="">Seleccione...</option>
                                                    <option value="1">Alta</option>
                                                    <option value="2">Media</option>
                                                    <option value="3">Baja</option>
                                                </select>
                                            </div>
                                        </div>
                                        <button type="button" className="btn btn-success w-100">
                                            Enviar Solicitud
                                        </button>
                                    </form>
                                </div>

                                <div className="card-header bg-secondary text-white border-top mt-2">
                                    <h5 className="card-title mb-0">Mis Solicitudes</h5>
                                </div>
                                <div className="card-body">
                                    <div className="table-responsive">
                                        <table className="table table-striped table-hover">
                                            <thead>
                                                <tr>
                                                    <th>ID</th>
                                                    <th>Título</th>
                                                    <th>Estado</th>
                                                    <th>Fecha</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {/* Filas dinámicas del backend */}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* VISTA OPERADOR */}
                        {rol === "Operador" && (
                            <div className="card mt-3 shadow-sm border-0">
                                <div className="card-body">
                                    <h5 className="text-success mb-3">Gestión de Solicitudes</h5>
                                    <div className="table-responsive">
                                        <table className="table table-hover align-middle">
                                            <thead>
                                                <tr>
                                                    <th>ID</th>
                                                    <th>Título</th>
                                                    <th>Estado Actual</th>
                                                    <th>Acción</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {/* Filas dinámicas del backend */}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        )}

                        <p className="mt-5 text-muted small">
                            <strong>idTokenClaims:</strong> {JSON.stringify(accounts[0].idTokenClaims)}
                        </p>
                    </>
                )}
                
                <button onClick={cerrarSesion} className="btn btn-danger mt-3">
                    Cerrar sesión
                </button>
            </AuthenticatedTemplate>
        </div>
    );
}

export default App;