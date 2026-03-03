import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import type { AuthenticationResult } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./config/authConfig";
import App from "./App";
import "./index.css";

// ── Initialize MSAL once at app startup ───────────────────────────────────────
const msalInstance = new PublicClientApplication(msalConfig);

// Set active account if one exists on page load
if (!msalInstance.getActiveAccount() && msalInstance.getAllAccounts().length > 0) {
    msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0]);
}

// Listen for successful login and set active account
msalInstance.addEventCallback((event) => {
    if (
        event.eventType === EventType.LOGIN_SUCCESS &&
        event.payload
    ) {
        const payload = event.payload as AuthenticationResult;
        msalInstance.setActiveAccount(payload.account);
    }
});

const container = document.getElementById("root");
if (!container) throw new Error("Root element #root not found in index.html");

createRoot(container).render(
    <StrictMode>
        <MsalProvider instance={msalInstance}>
            <App />
        </MsalProvider>
    </StrictMode>
);
