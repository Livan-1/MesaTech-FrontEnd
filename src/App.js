import logo from './logo.svg';
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest, apiRequest } from './authConfig';
import { useEffect, useState } from 'react';
import Axios from 'axios';


function App() {

    const { instance, accounts } = useMsal();
    const [usuarioBackend, setUsuarioBackend] = useState(null);
    const [errorBackend, setErrorBackend] = useState(null);


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
                <h2 style={{ color: "green" }} className="alert alert-success mt-3">Usuario autenticado con exito</h2>
                {accounts.length > 0 && (
                    <>
                        <div>
                            <p>
                                Nombre de usuario:
                                {"  "}
                                {accounts[0].name}
                            </p>
                            <p>
                                Correo electrónico:
                                {" "}
                                {accounts[0].username}
                            </p>
                            <p>
                                id:
                                {" "}
                                {accounts[0].idTokenClaims.oid}
                            </p>
                        </div>
                        <div class="card">
                            <div class="card-header">
                                <h5 class="card-title mt-1">Ingrese los datos de la solicitud</h5>
                                <div class="card-body">
                                    <form>
                                        <div class="form-group">
                                            <label for="asuntoInput">Asunto de la solicitud</label>
                                            <input type="text" class="form-control mt-2" id="asuntoInput" placeholder="Ingrese el asunto" />
                                            <button type="submit" class="btn btn-primary mt-2">Enviar solicitud</button>
                                        </div>
                                    </form>
                                </div>
                            </div>
                        </div>
                        <div class="card mt-3">
                            <div class="card-header">
                                <h5 class="card-title mt-1">Consultar solicitudes</h5>
                            </div>
                            <div class="card-body">
                                <button type="button" class="btn btn-primary mt-2">Consultar solicitudes</button>
                            </div>
                        </div>

                        <p class="mt-5">
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
