import logo from './logo.svg';
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';


function App() {

  const { instance, accounts } = useMsal();

  const iniciarSesion = () => {
    instance.loginRedirect(loginRequest)
      .catch(error => {
        console.error(error);
      });
  }

  const cerrarSesion = () => {
    instance.logoutRedirect();
  }

  return (
    <div style={{ padding: "30px" }}>

      <h1>Login con Microsoft Entra ID</h1>

      <UnauthenticatedTemplate>

        <p>
          El usuario no está autenticado.
        </p>

        <button onClick={iniciarSesion}>
          Iniciar sesión
        </button>

      </UnauthenticatedTemplate>


      <AuthenticatedTemplate>

        <h2>Usuario autenticado</h2>

        {accounts.length > 0 && (
          <>
            <p>
              Nombre:
              {" "}
              {accounts[0].name}
            </p>

            <p>
              Usuario:
              {" "}
              {accounts[0].username}
            </p>

            <p>
              id:
              {" "}
              {accounts[0].idTokenClaims.oid}
            </p>

            <p>
              idTokenClaims:
              {" "}
              {JSON.stringify(accounts[0].idTokenClaims)}
            </p>
          </>
        )}

        <button onClick={cerrarSesion}>
          Cerrar sesión
        </button>

      </AuthenticatedTemplate>

    </div>
  );
}

export default App;
