import { useState, useEffect } from "react";
import type { FC } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../config/authConfig";

// ─── SVG Icons ────────────────────────────────────────────────────────────────

const IconAlertCircle: FC<{ size?: number }> = ({ size = 16 }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
        stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <circle cx="12" cy="12" r="10" />
        <line x1="12" y1="8"  x2="12"    y2="12" />
        <line x1="12" y1="16" x2="12.01" y2="16" />
    </svg>
);

const MicrosoftLogo: FC = () => (
    <svg width="20" height="20" viewBox="0 0 21 21" fill="none">
        <rect x="1"  y="1"  width="9" height="9" fill="#f25022" />
        <rect x="11" y="1"  width="9" height="9" fill="#7fba00" />
        <rect x="1"  y="11" width="9" height="9" fill="#00a4ef" />
        <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
);

// ─── Component ────────────────────────────────────────────────────────────────

const LoginPage: FC = () => {
    const { instance }            = useMsal();
    const [mounted, setMounted]   = useState(false);
    const [loading, setLoading]   = useState(false);
    const [error, setError]       = useState<string | null>(null);

    useEffect(() => { setMounted(true); }, []);

    const handleAzureLogin = async (): Promise<void> => {
        setLoading(true);
        setError(null);
        try {
            await instance.loginPopup(loginRequest);
            // App.tsx useIsAuthenticated() will auto-update → Dashboard renders
        } catch (err) {
            if (err instanceof Error) {
                if (err.message.includes("user_cancelled") ||
                    err.message.includes("popup_window_error")) {
                    setError("Login was cancelled. Please try again.");
                } else if (err.message.includes("network")) {
                    setError("Network error. Check your internet connection.");
                } else {
                    setError(`Login failed: ${err.message}`);
                }
            } else {
                setError("An unexpected error occurred. Please try again.");
            }
        } finally {
            setLoading(false);
        }
    };

    return (
        <div style={{
            height: "100vh",
            width: "100%",
            display: "flex",
            background: "linear-gradient(135deg, #0a1628 0%, #0d2d5e 40%, #1565c0 100%)",
            fontFamily: "'Segoe UI', system-ui, sans-serif",
            position: "relative",
            overflow: "hidden",
        }}>

            {/* Background decorative circles */}
            <div style={{ position: "absolute", inset: 0, overflow: "hidden", pointerEvents: "none" }}>
                <div style={{ position: "absolute", width: 600, height: 600, borderRadius: "50%", border: "1px solid rgba(255,255,255,0.04)", top: -200, left: -200 }} />
                <div style={{ position: "absolute", width: 400, height: 400, borderRadius: "50%", border: "1px solid rgba(255,255,255,0.06)", top: -100, left: -100 }} />
                <div style={{ position: "absolute", width: 500, height: 500, borderRadius: "50%", border: "1px solid rgba(255,255,255,0.03)", bottom: -150, right: -150 }} />
                <div style={{ position: "absolute", width: 300, height: 300, borderRadius: "50%", background: "radial-gradient(circle, rgba(21,101,192,0.15) 0%, transparent 70%)", top: "20%", right: "10%" }} />
                <div style={{ position: "absolute", width: 200, height: 200, borderRadius: "50%", background: "radial-gradient(circle, rgba(66,165,245,0.1) 0%, transparent 70%)", bottom: "30%", left: "5%" }} />
            </div>

            {/* ── Left branding panel ─────────────────────────────────────── */}
            <div style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                padding: "60px 80px",
                position: "relative",
                opacity: mounted ? 1 : 0,
                transform: mounted ? "translateX(0)" : "translateX(-30px)",
                transition: "opacity 0.6s ease, transform 0.6s ease",
            }}>
                {/* Logo */}
                <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 60 }}>
                    <div style={{
                        width: 52, height: 52, borderRadius: "50%",
                        background: "linear-gradient(135deg, #42a5f5, #1565c0)",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        boxShadow: "0 4px 20px rgba(66,165,245,0.4)",
                    }}>
                        <span style={{ color: "#fff", fontWeight: 900, fontSize: 16 }}>HN</span>
                    </div>
                    <div>
                        <div style={{ color: "#fff", fontWeight: 800, fontSize: 18 }}>HEALTH NETWORK ONE</div>
                        <div style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, letterSpacing: "1px", textTransform: "uppercase" }}>Provider Web Portal</div>
                    </div>
                </div>

                {/* Hero */}
                <h1 style={{ color: "#fff", fontSize: 42, fontWeight: 800, lineHeight: 1.2, margin: "0 0 20px", letterSpacing: "-1px" }}>
                    Welcome to the<br />
                    <span style={{ background: "linear-gradient(90deg,#42a5f5,#90caf9)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
                        Provider Portal
                    </span>
                </h1>
                <p style={{ color: "rgba(255,255,255,0.6)", fontSize: 16, lineHeight: 1.7, maxWidth: 420, margin: "0 0 48px" }}>
                    Manage your claims, referrals, and member information all in one secure, streamlined dashboard.
                </p>

                {/* Feature list */}
                <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                    {[
                        { icon: "📋", text: "Submit and track claims in real time" },
                        { icon: "👥", text: "Access member eligibility instantly" },
                        { icon: "🔗", text: "Manage referrals with ease" },
                        { icon: "🔒", text: "Secure, HIPAA-compliant platform" },
                    ].map(({ icon, text }) => (
                        <div key={text} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                            <div style={{
                                width: 36, height: 36, borderRadius: "50%",
                                background: "rgba(255,255,255,0.08)",
                                display: "flex", alignItems: "center", justifyContent: "center",
                                fontSize: 16,
                            }}>{icon}</div>
                            <span style={{ color: "rgba(255,255,255,0.7)", fontSize: 14 }}>{text}</span>
                        </div>
                    ))}
                </div>
            </div>

            {/* ── Right login card ────────────────────────────────────────── */}
            <div style={{
                width: 480,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                padding: 40,
                opacity: mounted ? 1 : 0,
                transform: mounted ? "translateX(0)" : "translateX(30px)",
                transition: "opacity 0.6s ease 0.15s, transform 0.6s ease 0.15s",
            }}>
                <div style={{
                    width: "100%",
                    background: "#fff",
                    borderRadius: 16,
                    padding: 44,
                    boxShadow: "0 24px 64px rgba(0,0,0,0.35)",
                }}>

                    {/* Card header */}
                    <div style={{ marginBottom: 36 }}>
                        <h2 style={{ fontSize: 26, fontWeight: 800, color: "#0d2d5e", margin: "0 0 8px" }}>
                            Sign In
                        </h2>
                        <p style={{ fontSize: 14, color: "#7a8899", margin: 0 }}>
                            Use your organization account to access the portal
                        </p>
                    </div>

                    {/* Error message */}
                    {error !== null && (
                        <div style={{
                            background: "#ffebee", border: "1px solid #ef9a9a",
                            borderRadius: 8, padding: "12px 14px", marginBottom: 24,
                            display: "flex", alignItems: "center", gap: 8,
                            fontSize: 13, color: "#c62828",
                        }}>
                            <IconAlertCircle size={16} />
                            {error}
                        </div>
                    )}

                    {/* Azure AD Sign In button */}
                    <button
                        onClick={() => { void handleAzureLogin(); }}
                        disabled={loading}
                        style={{
                            width: "100%",
                            padding: "14px 20px",
                            background: loading ? "#90caf9" : "#0078d4",
                            color: "#fff",
                            border: "none",
                            borderRadius: 8,
                            fontSize: 15,
                            fontWeight: 600,
                            cursor: loading ? "not-allowed" : "pointer",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            gap: 12,
                            boxShadow: loading ? "none" : "0 4px 16px rgba(0,120,212,0.4)",
                            transition: "all 0.2s",
                            letterSpacing: "0.2px",
                        }}
                        onMouseEnter={(e) => {
                            if (!loading) {
                                (e.currentTarget as HTMLButtonElement).style.background = "#106ebe";
                                (e.currentTarget as HTMLButtonElement).style.transform  = "translateY(-1px)";
                                (e.currentTarget as HTMLButtonElement).style.boxShadow  = "0 6px 20px rgba(0,120,212,0.5)";
                            }
                        }}
                        onMouseLeave={(e) => {
                            if (!loading) {
                                (e.currentTarget as HTMLButtonElement).style.background = "#0078d4";
                                (e.currentTarget as HTMLButtonElement).style.transform  = "none";
                                (e.currentTarget as HTMLButtonElement).style.boxShadow  = "0 4px 16px rgba(0,120,212,0.4)";
                            }
                        }}
                    >
                        {loading ? (
                            <>
                                <div style={{
                                    width: 18, height: 18,
                                    border: "2px solid rgba(255,255,255,0.3)",
                                    borderTop: "2px solid #fff",
                                    borderRadius: "50%",
                                    animation: "spin 0.8s linear infinite",
                                }} />
                                Signing in…
                            </>
                        ) : (
                            <>
                                <MicrosoftLogo />
                                Sign in with Microsoft
                            </>
                        )}
                    </button>

                    <style>{`@keyframes spin { to { transform: rotate(360deg) } }`}</style>

                    {/* Divider */}
                    <div style={{ margin: "24px 0", display: "flex", alignItems: "center", gap: 12 }}>
                        <div style={{ flex: 1, height: 1, background: "#e8ecf0" }} />
                        <span style={{ fontSize: 11, color: "#aab4be", whiteSpace: "nowrap" }}>
                            Secured by Microsoft Azure AD
                        </span>
                        <div style={{ flex: 1, height: 1, background: "#e8ecf0" }} />
                    </div>

                    {/* Info box */}
                    <div style={{
                        background: "#f0f7ff",
                        border: "1px solid #cce4ff",
                        borderRadius: 8,
                        padding: "12px 14px",
                        fontSize: 12,
                        color: "#1565c0",
                        lineHeight: 1.6,
                    }}>
                        <strong>Health Network One employees:</strong><br />
                        Use your <code style={{ background: "#daeeff", padding: "1px 4px", borderRadius: 3 }}>@pecportaldev.onmicrosoft.com</code> account to sign in.
                    </div>

                    {/* Footer */}
                    <div style={{ marginTop: 28, paddingTop: 20, borderTop: "1px solid #f0f2f5", textAlign: "center" }}>
                        <p style={{ fontSize: 11, color: "#9aabb8", margin: 0, lineHeight: 1.6 }}>
                            Protected by enterprise-grade security<br />
                            © 2026 Health Network One. All rights reserved.
                        </p>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default LoginPage;
