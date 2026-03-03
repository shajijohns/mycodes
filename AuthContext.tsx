import { createContext, useContext } from "react";
import type { FC, ReactNode } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "../config/authConfig";

// ── Types ─────────────────────────────────────────────────────────────────────

interface AuthUser {
    username:  string;
    email:     string;
    role:      string;
    token:     string;
}

interface AuthContextValue {
    user:            AuthUser | null;
    isAuthenticated: boolean;
    login:           () => Promise<void>;
    logout:          () => void;
}

// ── Context ───────────────────────────────────────────────────────────────────

const AuthContext = createContext<AuthContextValue | null>(null);

// ── Provider ──────────────────────────────────────────────────────────────────

export const AuthProvider: FC<{ children: ReactNode }> = ({ children }) => {
    const { instance, accounts } = useMsal();
    const isAuthenticated        = useIsAuthenticated();

    const account = accounts[0] ?? null;

    const user: AuthUser | null = account
        ? {
            username: account.name     ?? account.username ?? "User",
            email:    account.username ?? "",
            role:     "Provider",
            token:    localStorage.getItem("authToken") ?? "",
        }
        : null;

    const login = async (): Promise<void> => {
        await instance.loginPopup(loginRequest);
    };

    const logout = (): void => {
        localStorage.clear();
        void instance.logoutPopup({
            postLogoutRedirectUri: window.location.origin,
        });
    };

    return (
        <AuthContext.Provider value={{ user, isAuthenticated, login, logout }}>
            {children}
        </AuthContext.Provider>
    );
};

// ── Hook ──────────────────────────────────────────────────────────────────────

export function useAuth(): AuthContextValue {
    const ctx = useContext(AuthContext);
    if (!ctx) throw new Error("useAuth must be used inside <AuthProvider>");
    return ctx;
}
