
import React, { useState } from "react";

function App() {
  const [file, setFile] = useState(null);
  const [jd, setJd] = useState("");
  const [company, setCompany] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");

  const getFilenameFromDisposition = (disposition) => {
    if (!disposition) return null;
    const match = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/i.exec(disposition);
    if (!match) return null;
    let name = match[1];
    if (name.startsWith('"') && name.endsWith('"')) name = name.slice(1, -1);
    return name;
  };

  const onSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setStatus("");

    if (!file) return setError("Please upload a DOCX resume file.");
    if (!jd.trim()) return setError("Please paste the job description.");

    try {
      setLoading(true);

      const form = new FormData();
      form.append("resume", file);
      form.append("jd", jd);
      form.append("company", company);

      const apiUrl = process.env.REACT_APP_API_URL || "http://localhost:4000";
      const res = await fetch(`${apiUrl}/api/modify-resume`, {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(text || "Failed to tailor resume.");
      }

      const blob = await res.blob();
      const disposition = res.headers.get("Content-Disposition");
      const filenameFromServer = getFilenameFromDisposition(disposition);

      const fallbackCompany = company && company.trim() ? company.trim() : "modified";
      const fallbackName = file.name.replace(/\.docx$/i, "") + " - " + fallbackCompany + ".docx";
      const filename = filenameFromServer || fallbackName;

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      setStatus("Downloaded tailored resume.");
    } catch (err) {
      console.error(err);
      setError(err.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: "100vh",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      background: "#0f172a",
      color: "#e5e7eb",
      fontFamily: "system-ui, -apple-system, Segoe UI, sans-serif",
      padding: 16
    }}>
      <div style={{
        width: "100%",
        maxWidth: 720,
        background: "#020617",
        borderRadius: 16,
        padding: 24,
        border: "1px solid rgba(148,163,184,0.35)",
        boxShadow: "0 20px 45px rgba(0,0,0,0.5)"
      }}>
        <h1 style={{ margin: 0, fontSize: 24 }}>Resume Tailor (v2.4)</h1>
        <p style={{ marginTop: 8, marginBottom: 20, color: "#9ca3af", fontSize: 14 }}>
          Upload a DOCX resume, paste the JD, enter the company name, and download the tailored DOCX.
          This version uses GPT to identify Experience/Skills lines to rewrite.
        </p>

        <form onSubmit={onSubmit}>
          <label style={{ display: "block", marginBottom: 6, fontSize: 14 }}>Resume (DOCX)</label>
          <input
            type="file"
            accept=".docx"
            onChange={(e) => setFile(e.target.files?.[0] || null)}
            style={{ width: "100%", marginBottom: 16 }}
          />

          <label style={{ display: "block", marginBottom: 6, fontSize: 14 }}>Company Name</label>
          <input
            type="text"
            value={company}
            onChange={(e) => setCompany(e.target.value)}
            placeholder="e.g. Vetcove"
            style={{ width: "100%", padding: 8, borderRadius: 8, border: "1px solid #4b5563", background: "#020617", color: "#e5e7eb", marginBottom: 16 }}
          />

          <label style={{ display: "block", marginBottom: 6, fontSize: 14 }}>Job Description</label>
          <textarea
            value={jd}
            onChange={(e) => setJd(e.target.value)}
            rows={10}
            placeholder="Paste the JD here..."
            style={{ width: "100%", padding: 8, borderRadius: 8, border: "1px solid #4b5563", background: "#020617", color: "#e5e7eb", resize: "vertical", marginBottom: 16 }}
          />

          {error ? (
            <div style={{ background: "rgba(239,68,68,0.12)", color: "#fecaca", padding: 10, borderRadius: 8, marginBottom: 12, fontSize: 13 }}>
              {error}
            </div>
          ) : status ? (
            <div style={{ background: "rgba(34,197,94,0.12)", color: "#bbf7d0", padding: 10, borderRadius: 8, marginBottom: 12, fontSize: 13 }}>
              {status}
            </div>
          ) : null}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: "100%",
              padding: "10px 12px",
              borderRadius: 999,
              border: "none",
              cursor: loading ? "default" : "pointer",
              fontWeight: 700,
              background: loading ? "rgba(59,130,246,0.55)" : "linear-gradient(135deg, #4f46e5, #3b82f6)",
              color: "#f9fafb",
              boxShadow: "0 10px 25px rgba(59,130,246,0.5)"
            }}
          >
            {loading ? "Tailoring..." : "Tailor & Download"}
          </button>
        </form>

        <div style={{ marginTop: 14, color: "#94a3b8", fontSize: 12 }}>
          Server must be running on <b>localhost:4000</b>.
        </div>
      </div>
    </div>
  );
}

export default App;
