import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const NAVY = "#1B3A5C";
const GOLD = "#C9A84C";

const SEQ_LABELS = ["", "Cold Intro", "Follow-Up", "Soft Check-In", "Last Try", "Pre-Quarterly"];

const SEQ_INSTRUCTIONS = {
  1: "Write Email 1 — a cold intro. First time reaching out. Be specific about what impressed Jesse about their business based on website info and notes. Warm, genuine, low-pressure. Mention his landscaping background and acquisition goals naturally. End with an invitation for coffee or lunch.",
  2: "Write Email 2 — a follow-up (sent ~1 week after email 1). No response yet. Keep it short — 2-3 paragraphs. Acknowledge they're busy. Restate interest briefly. Add a new angle or something from their website. Simple ask.",
  3: "Write Email 3 — a soft check-in (~2 weeks after email 2). Very short and casual. Light touch, no pressure. Just keeping the door open.",
  4: "Write Email 4 — a last try before going quiet. Direct but respectful. Acknowledge timing may be off. Leave the door open for the future.",
  5: "Write Email 5 — final touch before quarterly outreach. Brief, warm note. Tell them you'll stop reaching out regularly but will check in occasionally. No ask. Genuine goodbye for now.",
};

const SYSTEM_PROMPT = `You are writing outreach emails on behalf of Jesse Horine, owner of 10 Cent Investments.

Jesse's background (use naturally, don't cram all into every email):
- Been in landscaping since age 14
- Marine Corps Officer (veteran)
- MBA from CU Boulder
- Lives in Golden, CO
- Actively looking to acquire and operate a service business in the Denver area
- Email: Jesse@10CentInvestments.com | Phone: (803) 367-2677 | Website: www.10centinvestments.com

Jesse's writing style (match precisely):
- Warm, direct, genuine — never salesy or pushy
- Short paragraphs, conversational but professional
- Opens with something personal or specific to the business
- Low-pressure — wants to learn and connect, not just buy
- Closes with a simple specific ask (coffee, lunch, call)
- Signature: "Best," or "Very respectfully," then full name, title, phone, email, website
- Never uses buzzwords or corporate jargon
- Concise — 4 to 6 short paragraphs max

Example email (cold intro):
"Hey Mark,
I hope all is well and that the year is off to a strong start for you and the team at Elite Landscape.
My name is Jesse Horine. I've been in landscaping since I was 14, served as a Marine Corps Officer, and earned my MBA from CU Boulder. I now live in Golden and am actively looking to acquire and operate a landscaping company here in the Denver area.
I came across Elite Landscape & Outdoor Living and was really impressed with the quality of your hardscape and outdoor living work. If you're ever open to a conversation — whether now or in the future — I'd love to connect and learn more about what you've built.
I'd be happy to meet for a quick coffee or lunch sometime in the next few weeks if you're open to it.
Best,
Jesse Horine, Owner
M: (803) 367-2677
E: Jesse@10CentInvestments.com
www.10centinvestments.com"

Output ONLY valid JSON — no markdown, no extra text:
{"subject": "...", "body": "..."}`;

export default function App() {
  const [prospects, setProspects] = useState([]);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(null);
  const [seq, setSeq] = useState(1);
  const [extraNotes, setExtraNotes] = useState("");
  const [loading, setLoading] = useState(false);
  const [scraping, setScraping] = useState(false);
  const [scrapedInfo, setScrapedInfo] = useState("");
  const [output, setOutput] = useState(null);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const [fileLoaded, setFileLoaded] = useState(false);
  const fileRef = useRef();

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates: true });
        const ws = wb.Sheets["Businesses"];
        if (!ws) { setError("Could not find the Businesses sheet."); return; }
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let headerRow = -1;
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][0]).includes("Company")) { headerRow = i; break; }
        }
        if (headerRow === -1) { setError("Could not find header row."); return; }
        const headers = data[headerRow];
        const parsed = [];
        for (let i = headerRow + 1; i < data.length; i++) {
          const row = data[i];
          if (!row[0]) continue;
          const p = {};
          headers.forEach((h, idx) => { p[h] = row[idx] || ""; });
          p._company  = String(p["Company Name"] || "").trim();
          p._contact  = String(p["Owner / Contact"] || "").trim();
          p._industry = String(p["Industry"] || "").trim();
          p._location = String(p["Location"] || "").trim();
          p._email    = String(p["Email"] || "").trim();
          p._website  = String(p["Website"] || "").trim();
          p._status   = String(p["Lead Status"] || "").trim();
          p._value    = p["Est. Value ($)"] || "";
          p._notes    = String(p["Notes"] || "").trim();
          p._seq      = String(p["Sequence\nStage"] || p["Sequence Stage"] || "").trim();
          if (p._company) parsed.push(p);
        }
        setProspects(parsed);
        setFileLoaded(true);
        setError("");
      } catch (err) {
        setError("Error reading file: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  }

  async function scrapeWebsite(url) {
    if (!url) return;
    setScraping(true);
    setScrapedInfo("");
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          tools: [{ type: "web_search_20250305", name: "web_search" }],
          messages: [{
            role: "user",
            content: `Visit this business website and extract key information useful for a personalized acquisition outreach email: ${url}. Extract: what they do, specialties, notable projects, years in business, anything unique or impressive. Keep it to 3-5 bullet points.`
          }]
        })
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      const text = data.content.filter(b => b.type === "text").map(b => b.text).join("\n");
      setScrapedInfo(text);
    } catch (err) {
      setScrapedInfo("Could not scrape website: " + err.message);
    } finally {
      setScraping(false);
    }
  }

  function selectProspect(p) {
    setSelected(p);
    setOutput(null);
    setScrapedInfo("");
    setError("");
  }

  async function generate() {
    if (!selected) { setError("Please select a prospect first."); return; }
    setLoading(true);
    setError("");
    setOutput(null);

    const userPrompt = `Write a ${SEQ_LABELS[seq]} email (Email ${seq} of 5) for this prospect:

Company: ${selected._company}
Contact/Owner: ${selected._contact || "Unknown"}
Industry: ${selected._industry || "Unknown"}
Location: ${selected._location || "Unknown"}
Website: ${selected._website || "Unknown"}
Lead Status: ${selected._status || "Unknown"}
Estimated Value: ${selected._value ? "$" + Number(selected._value).toLocaleString() : "Unknown"}
CRM Notes: ${selected._notes || "None"}
${scrapedInfo ? `Website Research:\n${scrapedInfo}` : ""}
${extraNotes ? `Additional context from Jesse: ${extraNotes}` : ""}

Instructions: ${SEQ_INSTRUCTIONS[seq]}`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system: SYSTEM_PROMPT,
          messages: [{ role: "user", content: userPrompt }]
        })
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      const raw = data.content[0].text.replace(/```json|```/g, "").trim();
      setOutput(JSON.parse(raw));
    } catch (err) {
      setError("Error: " + err.message);
    } finally {
      setLoading(false);
    }
  }

  function copyEmail() {
    if (!output) return;
    navigator.clipboard.writeText(`Subject: ${output.subject}\n\n${output.body}`);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  function openOutlook() {
    if (!output || !selected) return;
    const to = encodeURIComponent(selected._email || "");
    const sub = encodeURIComponent(output.subject);
    const body = encodeURIComponent(output.body);
    window.location.href = `mailto:${to}?subject=${sub}&body=${body}`;
  }

  const filtered = prospects.filter(p =>
    p._company.toLowerCase().includes(search.toLowerCase()) ||
    p._contact.toLowerCase().includes(search.toLowerCase()) ||
    p._industry.toLowerCase().includes(search.toLowerCase())
  );

  const pillColor = (s) => ({
    Hot:    { bg: "#FFCDD2", color: "#C62828" },
    Warm:   { bg: "#FFE082", color: "#5D4037" },
    Cold:   { bg: "#BBDEFB", color: "#1565C0" },
    Closed: { bg: "#E0E0E0", color: "#616161" },
  }[s] || { bg: "#F5F5F5", color: "#999" });

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", background: "#FAF8F3", minHeight: "100vh" }}>

      {/* Header */}
      <div style={{ background: NAVY, borderBottom: `3px solid ${GOLD}`, padding: "0 32px", height: 60, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ color: "white", fontSize: 18, letterSpacing: 0.5 }}>
          10 Cent <span style={{ color: GOLD }}>Investments</span>
        </div>
        <div style={{ color: GOLD, fontSize: 11, letterSpacing: 3, opacity: 0.8, textTransform: "uppercase" }}>
          AI Email Generator
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "32px 24px", display: "grid", gridTemplateColumns: "360px 1fr", gap: 28, alignItems: "start" }}>

        {/* LEFT */}
        <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>

          {/* Step 1 */}
          <Panel num={1} title="Load CRM">
            {!fileLoaded ? (
              <div onClick={() => fileRef.current.click()} style={{ border: "2px dashed #e2ddd4", borderRadius: 4, padding: "28px 16px", textAlign: "center", cursor: "pointer", background: "#FAF8F3" }}>
                <div style={{ fontSize: 28, marginBottom: 8 }}>📊</div>
                <div style={{ fontSize: 13, color: "#555" }}><strong style={{ color: NAVY }}>Click to upload CRM</strong><br />10_Cent_CRM.xlsx</div>
              </div>
            ) : (
              <div style={{ background: "#EEF7EE", border: "1px solid #C8E6C9", borderRadius: 4, padding: "12px 14px", fontSize: 13, color: "#2E7D32", display: "flex", alignItems: "center", gap: 8 }}>
                ✅ <strong>{prospects.length} prospects loaded</strong>
                <span style={{ color: "#888", marginLeft: "auto", cursor: "pointer", fontSize: 11 }} onClick={() => { setFileLoaded(false); setProspects([]); setSelected(null); }}>Change</span>
              </div>
            )}
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />

            {fileLoaded && (
              <>
                <input
                  value={search} onChange={e => setSearch(e.target.value)}
                  placeholder="🔍  Search company or contact..."
                  style={{ width: "100%", marginTop: 12, padding: "9px 12px", border: "1px solid #e2ddd4", borderRadius: 4, fontSize: 13, outline: "none", boxSizing: "border-box" }}
                />
                <div style={{ maxHeight: 260, overflowY: "auto", border: "1px solid #e2ddd4", borderRadius: 4, marginTop: 8 }}>
                  {filtered.length === 0 && <div style={{ padding: 16, textAlign: "center", color: "#aaa", fontSize: 12 }}>No results</div>}
                  {filtered.map((p, i) => {
                    const pc = pillColor(p._status);
                    const isSel = selected?._company === p._company;
                    return (
                      <div key={i} onClick={() => selectProspect(p)} style={{ padding: "10px 14px", borderBottom: "1px solid #f0ede8", cursor: "pointer", background: isSel ? "#EEF3F8" : "white", borderLeft: isSel ? `3px solid ${NAVY}` : "3px solid transparent", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                        <div>
                          <div style={{ fontSize: 13, fontWeight: 500 }}>{p._company}</div>
                          <div style={{ fontSize: 11, color: "#aaa", marginTop: 2 }}>{p._contact || "—"} · {p._industry || "—"}</div>
                        </div>
                        {p._status && <span style={{ background: pc.bg, color: pc.color, fontSize: 9, padding: "2px 7px", borderRadius: 20, textTransform: "uppercase", whiteSpace: "nowrap" }}>{p._status}</span>}
                      </div>
                    );
                  })}
                </div>
              </>
            )}

            {selected && (
              <div style={{ background: "#FAF8F3", border: "1px solid #e2ddd4", borderRadius: 4, padding: 12, marginTop: 12, fontSize: 12 }}>
                {[["Company", selected._company], ["Contact", selected._contact], ["Industry", selected._industry], ["Location", selected._location], ["Email", selected._email], ["Website", selected._website], ["Status", selected._status], ["Notes", selected._notes]].filter(([, v]) => v).map(([l, v]) => (
                  <div key={l} style={{ display: "flex", gap: 8, marginBottom: 5 }}>
                    <span style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: 1, width: 58, flexShrink: 0 }}>{l}</span>
                    <span style={{ color: "#1a1a1a", lineHeight: 1.4 }}>{String(v)}</span>
                  </div>
                ))}
                {selected._website && (
                  <button onClick={() => scrapeWebsite(selected._website)} disabled={scraping} style={{ marginTop: 10, width: "100%", padding: "8px 12px", background: scraping ? "#ccc" : NAVY, color: "white", border: "none", borderRadius: 4, fontSize: 11, cursor: scraping ? "not-allowed" : "pointer" }}>
                    {scraping ? "⏳ Scraping..." : "🌐 Scrape Website for Context"}
                  </button>
                )}
                {scrapedInfo && (
                  <div style={{ marginTop: 10, background: "#E8F5E9", border: "1px solid #C8E6C9", borderRadius: 4, padding: 10, fontSize: 11, color: "#2E7D32", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>
                    <strong>✅ Website Intel</strong><br />{scrapedInfo}
                  </div>
                )}
              </div>
            )}
          </Panel>

          {/* Step 2 */}
          <Panel num={2} title="Configure Email">
            <div style={{ marginBottom: 14 }}>
              <Label>Email in Sequence</Label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {[1,2,3,4].map(n => (
                  <button key={n} onClick={() => setSeq(n)} style={{ padding: "10px 8px", border: `1px solid ${seq===n ? NAVY : "#e2ddd4"}`, borderRadius: 4, background: seq===n ? NAVY : "white", cursor: "pointer", textAlign: "center" }}>
                    <span style={{ fontSize: 10, display: "block", color: seq===n ? "#adf" : "#aaa" }}>Email {n}</span>
                    <span style={{ fontSize: 11, fontWeight: 500, color: seq===n ? "white" : "#333" }}>{SEQ_LABELS[n]}</span>
                  </button>
                ))}
                <button onClick={() => setSeq(5)} style={{ padding: "10px 8px", border: `1px solid ${seq===5 ? NAVY : "#e2ddd4"}`, borderRadius: 4, background: seq===5 ? NAVY : "white", cursor: "pointer", textAlign: "center", gridColumn: "span 2" }}>
                  <span style={{ fontSize: 10, display: "block", color: seq===5 ? "#adf" : "#aaa" }}>Email 5</span>
                  <span style={{ fontSize: 11, fontWeight: 500, color: seq===5 ? "white" : "#333" }}>Pre-Quarterly Follow-Up</span>
                </button>
              </div>
            </div>

            <div style={{ marginBottom: 14 }}>
              <Label>Extra Context (optional)</Label>
              <textarea value={extraNotes} onChange={e => setExtraNotes(e.target.value)} placeholder="Met at trade show, specific detail you noticed..." style={{ width: "100%", padding: "10px 12px", border: "1px solid #e2ddd4", borderRadius: 4, fontSize: 13, resize: "vertical", minHeight: 72, outline: "none", lineHeight: 1.5, boxSizing: "border-box" }} />
            </div>

            <button onClick={generate} disabled={loading || !selected} style={{ width: "100%", padding: 14, background: loading || !selected ? "#aaa" : NAVY, color: "white", border: "none", borderRadius: 4, fontSize: 12, letterSpacing: 2, textTransform: "uppercase", cursor: loading || !selected ? "not-allowed" : "pointer" }}>
              {loading ? "⏳ Generating..." : "Generate Email"}
            </button>

            {error && <div style={{ marginTop: 10, background: "#FFEBEE", border: "1px solid #FFCDD2", color: "#C62828", padding: "10px 12px", borderRadius: 4, fontSize: 12 }}>{error}</div>}
          </Panel>
        </div>

        {/* RIGHT */}
        <div style={{ position: "sticky", top: 24 }}>
          <Panel num={3} title="Generated Email">
            {!output ? (
              <div style={{ padding: "60px 20px", textAlign: "center", color: "#bbb" }}>
                <div style={{ fontSize: 40, marginBottom: 12 }}>✉️</div>
                <div style={{ fontSize: 13, lineHeight: 1.6 }}>Select a prospect, choose an email type,<br />and hit Generate.</div>
              </div>
            ) : (
              <div style={{ padding: 20 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, paddingBottom: 14, borderBottom: "1px solid #e2ddd4" }}>
                  <div style={{ fontSize: 11, color: "#aaa", textTransform: "uppercase", letterSpacing: 1 }}>
                    Draft for: <strong style={{ color: NAVY }}>{selected?._contact || selected?._company}</strong>
                  </div>
                  <span style={{ background: "#F5E9C8", color: GOLD, fontSize: 9, padding: "3px 8px", borderRadius: 20, textTransform: "uppercase", letterSpacing: 1 }}>Email {seq} — {SEQ_LABELS[seq]}</span>
                </div>

                <div style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>Subject Line</div>
                <div style={{ fontSize: 14, fontWeight: 500, color: NAVY, padding: "10px 12px", background: "#FAF8F3", borderLeft: `3px solid ${GOLD}`, marginBottom: 20 }}>{output.subject}</div>

                <div style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>Body</div>
                <textarea
                  value={output.body}
                  onChange={e => setOutput({ ...output, body: e.target.value })}
                  style={{ width: "100%", padding: 16, border: "1px solid #e2ddd4", borderRadius: 4, fontSize: 13, lineHeight: 1.75, color: "#1a1a1a", background: "#FAF8F3", resize: "vertical", minHeight: 340, outline: "none", boxSizing: "border-box" }}
                />

                <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
                  <Btn onClick={generate} disabled={loading}>↺ Regenerate</Btn>
                  <Btn onClick={copyEmail}>{copied ? "✓ Copied!" : "📋 Copy"}</Btn>
                  <Btn onClick={openOutlook} primary>Open in Outlook →</Btn>
                </div>
              </div>
            )}
          </Panel>
        </div>
      </div>
    </div>
  );
}

function Panel({ num, title, children }) {
  return (
    <div style={{ background: "white", border: "1px solid #e2ddd4", borderRadius: 4, overflow: "hidden" }}>
      <div style={{ background: NAVY, padding: "13px 18px", display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ width: 22, height: 22, background: GOLD, color: NAVY, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 600, flexShrink: 0 }}>{num}</div>
        <div style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: GOLD, fontWeight: 500 }}>{title}</div>
      </div>
      <div style={{ padding: 18 }}>{children}</div>
    </div>
  );
}

function Label({ children }) {
  return <div style={{ fontSize: 10, letterSpacing: 1.5, textTransform: "uppercase", color: NAVY, marginBottom: 6, fontWeight: 500 }}>{children}</div>;
}

function Btn({ children, onClick, primary, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{ flex: 1, padding: "10px 8px", borderRadius: 4, fontSize: 11, letterSpacing: 1, textTransform: "uppercase", cursor: disabled ? "not-allowed" : "pointer", border: primary ? `1px solid ${GOLD}` : "1px solid #e2ddd4", background: primary ? GOLD : "white", color: NAVY, fontWeight: primary ? 600 : 400, opacity: disabled ? 0.5 : 1 }}>
      {children}
    </button>
  );
}
