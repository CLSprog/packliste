import { useEffect, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { initMsal, login, getAccount } from "./auth";
import PackingApp from "./PackingApp";

export default function App() {
  const [status, setStatus] = useState<"initializing" | "signed-out" | "signing-in" | "signed-in" | "error">("initializing");
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    initMsal()
      .then(() => {
        const existing = getAccount();
        if (existing) {
          setAccount(existing);
          setStatus("signed-in");
        } else {
          setStatus("signed-out");
        }
      })
      .catch((error) => {
        setErrorMessage(error instanceof Error ? error.message : "Unbekannter Fehler beim Start.");
        setStatus("error");
      });
  }, []);

  async function handleLogin() {
    setStatus("signing-in");
    setErrorMessage("");
    try {
      const signedInAccount = await login();
      setAccount(signedInAccount);
      setStatus("signed-in");
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "Anmeldung fehlgeschlagen.");
      setStatus("signed-out");
    }
  }

  if (status === "initializing") {
    return <div className="auth-screen"><p>Wird geladen …</p></div>;
  }

  if (status === "signed-in" && account) {
    return <PackingApp account={account} />;
  }

  return (
    <div className="auth-screen">
      <div className="auth-card">
        <h1>P03 Packliste</h1>
        <p>Bitte melde dich mit deinem Microsoft-Konto an, um deine Packlisten aus OneDrive zu laden.</p>
        {errorMessage && <p className="auth-error">{errorMessage}</p>}
        <button className="primary compact-action" disabled={status === "signing-in"} onClick={handleLogin}>
          {status === "signing-in" ? "Anmeldung läuft …" : "Mit Microsoft anmelden"}
        </button>
      </div>
    </div>
  );
}
