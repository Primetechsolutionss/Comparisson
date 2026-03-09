import { useState, useRef, useCallback } from "react";

const API_BASE = typeof window !== 'undefined' && window.location.hostname === 'localhost'
  ? 'http://localhost:8000'
  : (process.env.NEXT_PUBLIC_API_BASE || '');

const DEFAULT_ALLOWLIST = ['.pdf', '.docx', '.xlsx', '.dwg', '.dgn', '.ifc', '.zip', '.rvt', '.jpg', '.xls', '.doc', '.pptx'];

function FileDropZone({ label, description, file, files, onFile, onFiles, accept, icon, multiple }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);

  const hasContent = multiple ? (files && files.length > 0) : !!file;

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    if (multiple) {
      const dropped = Array.from(e.dataTransfer.files);
      if (dropped.length > 0) onFiles(dropped);
    } else {
      const f = e.dataTransfer.files[0];
      if (f) onFile(f);
    }
  }, [onFile, onFiles, multiple]);

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      style={{
        flex: 1,
        minHeight: 200,
        border: `2px dashed ${dragOver ? '#3B82F6' : hasContent ? '#22C55E' : '#334155'}`,
        borderRadius: 16,
        padding: 32,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        gap: 12,
        cursor: 'pointer',
        background: dragOver ? 'rgba(59,130,246,0.06)' : hasContent ? 'rgba(34,197,94,0.04)' : 'rgba(15,23,42,0.3)',
        transition: 'all 0.2s ease',
        position: 'relative',
        overflow: 'hidden',
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        multiple={!!multiple}
        onChange={(e) => {
          if (multiple) {
            const picked = Array.from(e.target.files);
            if (picked.length > 0) onFiles(picked);
          } else {
            if (e.target.files[0]) onFile(e.target.files[0]);
          }
        }}
        style={{ display: 'none' }}
      />
      <div style={{ fontSize: 40, opacity: 0.8 }}>{icon}</div>
      <div style={{ fontSize: 15, fontWeight: 600, color: '#E2E8F0', letterSpacing: '0.02em' }}>{label}</div>
      <div style={{ fontSize: 13, color: '#94A3B8', textAlign: 'center', maxWidth: 240, lineHeight: 1.5 }}>{description}</div>
      {!multiple && file && (
        <div style={{
          marginTop: 8, padding: '8px 16px',
          background: 'rgba(34,197,94,0.12)', borderRadius: 8,
          fontSize: 13, color: '#22C55E', fontWeight: 500,
          display: 'flex', alignItems: 'center', gap: 6,
        }}>
          <span>✓</span> {file.name}
        </div>
      )}
      {multiple && files && files.length > 0 && (
        <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 4, width: '100%', maxWidth: 300 }}>
          {files.map((f, i) => (
            <div key={i} style={{
              padding: '6px 12px',
              background: 'rgba(34,197,94,0.12)', borderRadius: 6,
              fontSize: 12, color: '#22C55E', fontWeight: 500,
              display: 'flex', alignItems: 'center', gap: 6,
              overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
            }}>
              <span>✓</span> {f.name}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, color, sub }) {
  return (
    <div style={{
      padding: '20px 24px',
      background: 'rgba(15,23,42,0.5)',
      borderRadius: 12,
      border: '1px solid #1E293B',
      minWidth: 140,
    }}>
      <div style={{ fontSize: 28, fontWeight: 700, color: color || '#E2E8F0', fontFamily: "'JetBrains Mono', monospace" }}>{value}</div>
      <div style={{ fontSize: 12, color: '#94A3B8', marginTop: 4, textTransform: 'uppercase', letterSpacing: '0.08em', fontWeight: 500 }}>{label}</div>
      {sub && <div style={{ fontSize: 11, color: '#64748B', marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function AllowlistEditor({ allowlist, onChange }) {
  const [open, setOpen] = useState(false);

  return (
    <div style={{ marginTop: 16 }}>
      <button
        onClick={() => setOpen(!open)}
        style={{
          background: 'none',
          border: 'none',
          color: '#64748B',
          fontSize: 13,
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          padding: '4px 0',
        }}
      >
        <span style={{ transition: 'transform 0.2s', transform: open ? 'rotate(90deg)' : 'none', display: 'inline-block' }}>▶</span>
        Allowed file types ({allowlist.length})
      </button>
      {open && (
        <div style={{
          marginTop: 8,
          padding: 16,
          background: 'rgba(15,23,42,0.4)',
          borderRadius: 10,
          display: 'flex',
          flexWrap: 'wrap',
          gap: 8,
        }}>
          {DEFAULT_ALLOWLIST.map(ext => (
            <label key={ext} style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              padding: '6px 12px',
              borderRadius: 6,
              background: allowlist.includes(ext) ? 'rgba(59,130,246,0.15)' : 'rgba(30,41,59,0.5)',
              border: `1px solid ${allowlist.includes(ext) ? '#3B82F6' : '#334155'}`,
              cursor: 'pointer',
              fontSize: 13,
              color: allowlist.includes(ext) ? '#93C5FD' : '#64748B',
              transition: 'all 0.15s ease',
              userSelect: 'none',
            }}>
              <input
                type="checkbox"
                checked={allowlist.includes(ext)}
                onChange={(e) => {
                  if (e.target.checked) onChange([...allowlist, ext]);
                  else onChange(allowlist.filter(x => x !== ext));
                }}
                style={{ display: 'none' }}
              />
              <span style={{
                width: 16, height: 16, borderRadius: 4,
                border: `2px solid ${allowlist.includes(ext) ? '#3B82F6' : '#475569'}`,
                background: allowlist.includes(ext) ? '#3B82F6' : 'transparent',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 10, color: '#fff', fontWeight: 700,
              }}>{allowlist.includes(ext) ? '✓' : ''}</span>
              {ext}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [masterFile, setMasterFile] = useState(null);
  const [deliveryFiles, setDeliveryFiles] = useState([]);
  const [allowlist, setAllowlist] = useState([...DEFAULT_ALLOWLIST]);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState('');
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');

  const canCompare = masterFile && deliveryFiles.length > 0 && !loading;

  const runComparison = async () => {
    setLoading(true);
    setError('');
    setResult(null);
    setProgress('Uploading files...');

    try {
      const formData = new FormData();
      formData.append('master_file', masterFile);
      deliveryFiles.forEach(f => formData.append('delivery_files', f));
      formData.append('allowlist', allowlist.join(','));

      setProgress('Running comparison engine...');

      const response = await fetch(`${API_BASE}/api/compare`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Comparison failed');
      }

      setProgress('Generating report...');
      const data = await response.json();
      setResult(data);
      setProgress('');
    } catch (err) {
      setError(err.message);
      setProgress('');
    } finally {
      setLoading(false);
    }
  };

  const reset = () => {
    setMasterFile(null);
    setDeliveryFiles([]);
    setResult(null);
    setError('');
    setProgress('');
  };

  const matchRateColor = result?.stats?.match_rate >= 99 ? '#22C55E' : result?.stats?.match_rate >= 95 ? '#EAB308' : '#EF4444';

  return (
    <div style={{
      minHeight: '100vh',
      background: '#0B1120',
      color: '#E2E8F0',
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, sans-serif",
    }}>
      {/* Header */}
      <header style={{
        padding: '24px 40px',
        borderBottom: '1px solid #1E293B',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 36, height: 36, borderRadius: 10,
            background: 'linear-gradient(135deg, #3B82F6, #8B5CF6)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: 18, fontWeight: 700,
          }}>L</div>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, letterSpacing: '-0.02em' }}>Leveransplan Validator</div>
            <div style={{ fontSize: 11, color: '#64748B', letterSpacing: '0.04em' }}>DOCUMENT DELIVERY CONTROL</div>
          </div>
        </div>
        <div style={{ fontSize: 12, color: '#475569' }}>
          Primetech Solutions
        </div>
      </header>

      {/* Main content */}
      <main style={{ maxWidth: 960, margin: '0 auto', padding: '48px 24px' }}>
        {!result ? (
          <>
            {/* Upload section */}
            <div style={{ textAlign: 'center', marginBottom: 48 }}>
              <h1 style={{
                fontSize: 32, fontWeight: 700, letterSpacing: '-0.03em',
                background: 'linear-gradient(135deg, #E2E8F0, #94A3B8)',
                WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent',
                marginBottom: 12,
              }}>
                Compare Delivery Against Master
              </h1>
              <p style={{ color: '#64748B', fontSize: 15, maxWidth: 520, margin: '0 auto', lineHeight: 1.6 }}>
                Upload your Leveransplan master file and a delivery sheet to validate document completeness.
              </p>
            </div>

            <div style={{ display: 'flex', gap: 20, marginBottom: 24 }}>
              <FileDropZone
                label="Master Leveransplan"
                description="The full master file with all packages and expected documents"
                file={masterFile}
                onFile={setMasterFile}
                accept=".xlsx,.xls"
                icon="📋"
              />
              <FileDropZone
                label="Delivery Sheet(s)"
                description="One or more delivery packages to check against the master"
                files={deliveryFiles}
                onFiles={setDeliveryFiles}
                accept=".xlsx,.xls"
                icon="📦"
                multiple
              />
            </div>

            <AllowlistEditor allowlist={allowlist} onChange={setAllowlist} />

            {/* Compare button */}
            <div style={{ marginTop: 32, display: 'flex', justifyContent: 'center' }}>
              <button
                onClick={runComparison}
                disabled={!canCompare}
                style={{
                  padding: '14px 48px',
                  fontSize: 15,
                  fontWeight: 600,
                  letterSpacing: '0.02em',
                  borderRadius: 12,
                  border: 'none',
                  cursor: canCompare ? 'pointer' : 'not-allowed',
                  background: canCompare
                    ? 'linear-gradient(135deg, #3B82F6, #6366F1)'
                    : '#1E293B',
                  color: canCompare ? '#fff' : '#475569',
                  transition: 'all 0.2s ease',
                  boxShadow: canCompare ? '0 4px 24px rgba(59,130,246,0.3)' : 'none',
                  position: 'relative',
                  overflow: 'hidden',
                }}
              >
                {loading ? (
                  <span style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span style={{
                      width: 18, height: 18, border: '2px solid rgba(255,255,255,0.3)',
                      borderTopColor: '#fff', borderRadius: '50%',
                      animation: 'spin 0.8s linear infinite',
                    }} />
                    {progress}
                  </span>
                ) : 'Compare'}
              </button>
            </div>

            {error && (
              <div style={{
                marginTop: 24, padding: '16px 20px',
                background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)',
                borderRadius: 10, color: '#FCA5A5', fontSize: 14, textAlign: 'center',
              }}>
                {error}
              </div>
            )}
          </>
        ) : (
          <>
            {/* Results */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 32 }}>
              <div>
                <h2 style={{ fontSize: 24, fontWeight: 700, letterSpacing: '-0.02em', margin: 0 }}>Comparison Results</h2>
                <p style={{ color: '#64748B', fontSize: 13, marginTop: 4 }}>
                  {result?.multi
                    ? `${result.stats.length} delivery sheets vs Master Leveransplan`
                    : `${deliveryFiles[0]?.name} vs Master Leveransplan`}
                </p>
              </div>
              <button onClick={reset} style={{
                padding: '10px 20px', borderRadius: 8, border: '1px solid #334155',
                background: 'transparent', color: '#94A3B8', cursor: 'pointer', fontSize: 13,
              }}>
                ← New Comparison
              </button>
            </div>

            {/* Stats grid — single delivery */}
            {!result.multi && (
              <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', marginBottom: 32 }}>
                <StatCard label="Match Rate" value={`${result.stats.match_rate.toFixed(1)}%`} color={matchRateColor} />
                <StatCard label="Raw Entries" value={result.stats.raw_row_count} sub="Delivery sheet rows" />
                <StatCard label="Compared" value={result.stats.unique_files_for_comparison} sub="After filtering" />
                <StatCard label="Found" value={result.stats.found} color="#22C55E" />
                <StatCard label="Not Found" value={result.stats.not_found} color={result.stats.not_found > 0 ? '#EF4444' : '#22C55E'} />
                <StatCard label="Revision Match" value={result.stats.revision_match} color={result.stats.revision_match > 0 ? '#EAB308' : '#94A3B8'} />
                <StatCard label="Anomalies" value={result.stats.flagged} color={result.stats.flagged > 0 ? '#EAB308' : '#94A3B8'} />
              </div>
            )}

            {/* Stats table — multiple deliveries */}
            {result.multi && (
              <div style={{ marginBottom: 32, overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#1B2A4A' }}>
                      {['Delivery Sheet', 'Compared', 'Found', 'Not Found', 'Match Rate', 'Revisions', 'Anomalies'].map(h => (
                        <th key={h} style={{ padding: '10px 14px', textAlign: 'left', color: '#fff', fontWeight: 600, whiteSpace: 'nowrap', border: '1px solid #334155' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {result.stats.map((row, i) => {
                      const s = row.stats;
                      const rateColor = s.match_rate >= 99 ? '#22C55E' : s.match_rate >= 95 ? '#EAB308' : '#EF4444';
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? 'rgba(15,23,42,0.4)' : 'rgba(15,23,42,0.2)' }}>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: '#E2E8F0', maxWidth: 240, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{row.delivery}</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: '#94A3B8' }}>{s.unique_files_for_comparison}</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: '#22C55E' }}>{s.found}</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: s.not_found > 0 ? '#EF4444' : '#22C55E' }}>{s.not_found}</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: rateColor, fontWeight: 700 }}>{s.match_rate.toFixed(1)}%</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: s.revision_match > 0 ? '#EAB308' : '#94A3B8' }}>{s.revision_match}</td>
                          <td style={{ padding: '9px 14px', border: '1px solid #1E293B', color: s.flagged > 0 ? '#EAB308' : '#94A3B8' }}>{s.flagged}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* Download button */}
            <a
              href={`${API_BASE}${result.download_url}`}
              download
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 10,
                padding: '14px 32px',
                fontSize: 15,
                fontWeight: 600,
                borderRadius: 12,
                border: 'none',
                background: 'linear-gradient(135deg, #22C55E, #16A34A)',
                color: '#fff',
                textDecoration: 'none',
                boxShadow: '0 4px 24px rgba(34,197,94,0.3)',
                transition: 'all 0.2s ease',
                cursor: 'pointer',
              }}
            >
              <span style={{ fontSize: 20 }}>📥</span> Download Excel Report
            </a>

            {/* Summary text */}
            <div style={{
              marginTop: 32,
              padding: 24,
              background: 'rgba(15,23,42,0.5)',
              borderRadius: 12,
              border: '1px solid #1E293B',
            }}>
              <h3 style={{ fontSize: 14, fontWeight: 600, color: '#94A3B8', textTransform: 'uppercase', letterSpacing: '0.06em', marginTop: 0, marginBottom: 16 }}>
                Detailed Summary
              </h3>
              <pre style={{
                fontSize: 13,
                lineHeight: 1.7,
                color: '#CBD5E1',
                fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                margin: 0,
              }}>
                {result.summary}
              </pre>
            </div>
          </>
        )}
      </main>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
        @keyframes spin { to { transform: rotate(360deg); } }
        * { box-sizing: border-box; margin: 0; }
        button:hover { opacity: 0.92; }
        a:hover { opacity: 0.92; }
      `}</style>
    </div>
  );
}
