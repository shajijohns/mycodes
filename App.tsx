import { useEffect } from "react";
import type { FC } from "react";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { AuthProvider } from "./context/AuthContext";
import LoginPage from "./components/LoginPage";
import Dashboard from "./components/Dashboard";

// ── Inner router — reads MSAL auth state ──────────────────────────────────────
const AppRouter: FC = () => {
    const isAuthenticated            = useIsAuthenticated();
    const { accounts, inProgress }   = useMsal();

    // Sync MSAL account info into localStorage for Dashboard to use
    useEffect(() => {
        if (accounts.length > 0) {
            const account = accounts[0];
            localStorage.setItem("username",  account.name     ?? account.username);
            localStorage.setItem("userEmail", account.username ?? "");
            localStorage.setItem("userRole",  "Provider");
        }
    }, [accounts]);

    // Show loading spinner while MSAL is initializing
    if (inProgress === InteractionStatus.Startup ||
        inProgress === InteractionStatus.HandleRedirect) {
        return (
            <div style={{
                height: "100vh", width: "100%",
                display: "flex", alignItems: "center", justifyContent: "center",
                background: "linear-gradient(135deg, #0a1628 0%, #0d2d5e 40%, #1565c0 100%)",
            }}>
                <div style={{ textAlign: "center", color: "#fff" }}>
                    <div style={{
                        width: 48, height: 48, margin: "0 auto 16px",
                        border: "4px solid rgba(255,255,255,0.2)",
                        borderTop: "4px solid #fff",
                        borderRadius: "50%",
                        animation: "spin 0.8s linear infinite",
                    }} />
                    <p style={{ fontSize: 14, opacity: 0.7 }}>Initializing…</p>
                    <style>{`@keyframes spin { to { transform: rotate(360deg) } }`}</style>
                </div>
            </div>
        );
    }

    if (isAuthenticated) {
        return <Dashboard onLogout={() => { localStorage.clear(); }} />;
    }

    return <LoginPage />;
};

// ── Root App — AuthProvider wraps everything ──────────────────────────────────
const App: FC = () => (
    <AuthProvider>
        <AppRouter />
    </AuthProvider>
);

export default App;
