import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest, apiRequest } from './authConfig';
import { useEffect, useState } from 'react';
import Axios from 'axios';

// Por si acaso cabros estas variables si estan estudiando el codigo pueden cambiarlas por sus correos pa probar por si acaso
const rolesPorCorreo = {
    "l.sepulveda.aulaeduca@gmail.com": "Cliente",
    "livansepulveda087@gmail.com": "Operador",
    "liv.sepulveda@duocuc.cl": "Administrador",
};

const obtenerRol = (correo) => {
    return rolesPorCorreo[correo] || "Desconocido";
}

function App() {

    const { instance, accounts } = useMsal();
    const [usuarioBackend, setUsuarioBackend] = useState(null);
    const [errorBackend, setErrorBackend] = useState(null);
    const correo = accounts[0]?.username;
    const rol = obtenerRol(correo);


    const iniciarSesion = () => {
        instance.loginRedirect(loginRequest)
            .catch(error => {
                console.error(error);
            });
    }

    const cerrarSesion = () => {
        instance.logoutRedirect();
    }

    useEffect(() => {
        if (accounts.length === 0) {
            return;
        }

        const obtenerUsuarioBackend = async () => {
            try {
                // solicitar a Entra ID un access token
                const tokenResponse = await instance.acquireTokenSilent({ ...apiRequest, account: accounts[0] });
                const accessToken = tokenResponse.accessToken;

                console.log(accessToken);

                // consumir servicio ahora que tenemos access token
                Axios.get("http://localhost:8080/api/usuario", { headers: { Authorization: `Bearer ${accessToken}` } })
                    .then((response) => {
                        console.log(response.data);
                        setUsuarioBackend(response.data);
                    })
                    .catch(
                        (error) => {
                            console.log(error);
                            setErrorBackend("Error consultando api");
                        }
                    )
            } catch (error) {
                console.log("Error obteniendo datos", error);
                setErrorBackend("No fue posible obtener el access token");
            }
        }

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
                <h2 className="alert alert-success mt-3">Usuario autenticado con exito</h2>
                {accounts.length > 0 && (
                    <>
                        <div>
                            <p>
                                Nombre de usuario:
                                {"  "}
                                {accounts[0].name}
                                {"  "}
                                {correo}
                            </p>
                            <p>
                                Rol:
                                {" "}
                                {rol}
                            </p>
                        </div>

                        {rol === "Administrador" && (
                            <div className="card mt-3">
                                <div>
                                    <h2 style={{ color: "green" }} className="alert alert-success mt-3">Bienvenido Administrador</h2>
                                </div>
                                <div className="card-header">
                                    <h5 className="card-title mt-1">Catalogo de categorias/prioridades</h5>
                                </div>
                                <div className="card-header">
                                    <h5 className="card-title mt-1">Solicitudes globales</h5>
                                </div>
                                <div className="card-body">
                                    <button type="button" className="btn btn-primary mt-2">
                                        Consultar solicitudes
                                    </button>
                                </div>
                            </div>
                        )}
                        {rol == "Cliente" && (
                            <div className="card mt-3">
                                <div className="card-header">
                                    <h5 className="card-title mt-1">Ingrese los datos de la solicitud</h5>
                                </div>
                                <div className="card-body">
                                    <form></form>
                                </div>
                            </div>
                        )}
                        {rol == "Operador" && (
                            <div className="card mt-3">
                                <div className="card-header">
                                    <h5 className="card-title mt-1">Consultar solicitudes asignadas</h5>
                                    <div className="card-body">
                                        <button type="button" className="btn btn-primary mt-2">Consultar solicitudes
                                        </button>
                                    </div>
                                </div>
                            </div>
                        )}

                        <p className="mt-5">
                            idTokenClaims:
                            {" "}
                            {JSON.stringify(accounts[0].idTokenClaims)}
                        </p>
                    </>
                )}
                <button onClick={cerrarSesion} className="btn btn-danger">
                    Cerrar sesión
                </button>
            </AuthenticatedTemplate>

        </div>
    );
}

export default App;
