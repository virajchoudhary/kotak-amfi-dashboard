import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Download, FileUp, Moon, RefreshCw, Sun } from 'lucide-react';
import {
    CartesianGrid,
    Legend,
    Line,
    LineChart,
    ResponsiveContainer,
    Tooltip,
    XAxis,
    YAxis,
} from 'recharts';

const API = import.meta.env.VITE_API_URL || window.location.origin;

const tabs = [
    { id: 'overview', label: 'Overview' },
    { id: 'sip', label: 'SIP Analysis' },
    { id: 'flat', label: 'Fund Flows (Flat)' },
    { id: 'form', label: 'AMFI Form Layout' },
    { id: 'archives', label: 'Archives' },
];

const sheetNames = {
    sip: 'AMFI-SIP',
    flatPrefix: "AMFI-Mar'25 to ",
    formSuffix: '-AMFI form',
};

function formatNumber(value) {
    if (typeof value !== 'number') return value ?? '';
    return new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(value);
}

function compactCrore(value) {
    if (!value) return '0';
    return new Intl.NumberFormat('en-IN', { maximumFractionDigits: 1 }).format(value);
}

function isBlankCell(value) {
    return value === null || value === undefined || String(value).trim() === '';
}

function isBlankRow(row = []) {
    return row.every(isBlankCell);
}

function columnLabel(value, index) {
    return isBlankCell(value) ? `Column ${index + 1}` : String(value).trim();
}

function prepareTable(sheet, limit) {
    const workbookHeaderIndex = 2;
    const allRows = sheet?.rows || [];
    const header = sheet?.columns?.length ? sheet.columns : allRows[workbookHeaderIndex] || [];
    const dataRows = allRows.slice(workbookHeaderIndex + 1).filter(row => !isBlankRow(row));
    const visibleRows = limit ? dataRows.slice(0, limit) : dataRows;
    const scanRows = [header, ...dataRows];
    const maxColumns = Math.max(sheet?.maxColumn || 0, ...scanRows.map(row => row.length));
    const indexes = Array.from({ length: maxColumns }, (_, index) => index)
        .filter(index => scanRows.some(row => !isBlankCell(row[index])));

    return {
        columns: indexes.map(index => ({ index, label: columnLabel(header[index], index) })),
        rows: visibleRows.map(row => indexes.map(index => row[index] ?? '')),
    };
}

function findSheet(data, predicate) {
    return Object.entries(data?.sheets || {}).find(([name]) => predicate(name))?.[1];
}

function sipSheet(data) {
    return data?.sheets?.[sheetNames.sip];
}

function flatSheet(data) {
    return findSheet(data, name => name.startsWith("AMFI-Mar") && !name.endsWith(sheetNames.formSuffix));
}

function formSheet(data) {
    return findSheet(data, name => name.startsWith("AMFI-Mar") && name.endsWith(sheetNames.formSuffix));
}

function Section({ title, subtitle, headerAction, children, id }) {
    return (
        <section className="card" id={id}>
            <div className="card-header">
                <div>
                    <h2>{title}</h2>
                    {subtitle && <p className="card-subtitle">{subtitle}</p>}
                </div>
                {headerAction && <div className="card-header-action">{headerAction}</div>}
            </div>
            <div className="card-body">{children}</div>
        </section>
    );
}

function DataTable({ sheet, limit }) {
    if (!sheet) return <div className="empty-state">No data loaded.</div>;
    const { columns, rows } = prepareTable(sheet, limit);
    if (!columns.length || !rows.length) return <div className="empty-state">No readable rows found.</div>;

    return (
        <div className="table-scroll">
            <table className="theory-table">
                <thead>
                    <tr>
                        {columns.map(column => (
                            <th key={column.index} title={column.label}>{column.label}</th>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {rows.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                            {row.map((cell, cellIndex) => (
                                <td key={cellIndex} className={typeof cell === 'number' ? 'numeric-cell' : undefined}>
                                    {formatNumber(cell)}
                                </td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}

function CustomTooltip({ active, payload, label }) {
    if (!active || !payload?.length) return null;
    return (
        <div className="chart-tooltip">
            <strong>{label}</strong>
            {payload.map(item => (
                <span key={item.dataKey}>{item.name}: {formatNumber(item.value)}</span>
            ))}
        </div>
    );
}

function Overview({ data, loading, onUpload, onRefresh, selectedFY = '' }) {
    const inputRef = useRef(null);
    const [fileName, setFileName] = useState('');
    const summary = data?.summary || {};
    const series = summary.timeSeries || [];

    async function submitUpload() {
        const file = inputRef.current?.files?.[0];
        if (!file) return;
        setFileName(file.name);
        await onUpload(file);
        if (inputRef.current) inputRef.current.value = "";
        setFileName("");
    }

    return (
        <>
            <Section 
                title="AMFI Ingestion" 
                subtitle="Upload monthly data sheets and trigger rollover baselines."
            >
                <div className="upload-row">
                    <label className="file-input">
                        <FileUp size={18} />
                        <span>{fileName || 'Select AMFI workbook'}</span>
                        <input
                            ref={inputRef}
                            type="file"
                            accept=".xlsx,.xls"
                            onChange={event => setFileName(event.target.files?.[0]?.name || '')}
                        />
                    </label>
                    <button className="btn-primary" onClick={submitUpload} disabled={loading || !fileName}>
                        {loading ? <span className="spinner" /> : <FileUp size={18} />}
                        Upload Monthly AMFI Report
                    </button>
                    <button className="btn-sm" onClick={() => onRefresh(selectedFY)} disabled={loading} title="Refresh data">
                        <RefreshCw size={16} />
                    </button>
                </div>

                {data?.warnings?.length ? (
                    <div className="warning-list">
                        {data.warnings.map((warning, index) => <span key={index}>{warning}</span>)}
                    </div>
                ) : null}

                <div className="kpi-grid">
                    <div className="kpi">
                        <span>Latest Month</span>
                        <strong>{summary.latestMonth || '—'}</strong>
                    </div>
                    <div className="kpi">
                        <span>Funds Mobilized</span>
                        <strong>{compactCrore(summary.latestFundsMobilized)}</strong>
                    </div>
                    <div className="kpi">
                        <span>Net Inflow</span>
                        <strong>{compactCrore(summary.latestNetInflow)}</strong>
                    </div>
                    <div className="kpi">
                        <span>Net AUM</span>
                        <strong>{compactCrore(summary.latestNetAum)}</strong>
                    </div>
                </div>
            </Section>

            <Section title="Time-Series Flow" subtitle="Aggregated monthly metrics from the flat master sheet.">
                <div className="chart-frame">
                    <ResponsiveContainer width="100%" height={360}>
                        <LineChart data={series} margin={{ top: 10, right: 40, left: 0, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                            <XAxis dataKey="month" stroke="var(--chart-axis)" tickMargin={10} interval={0} height={44} />
                            <YAxis stroke="var(--chart-axis)" tickMargin={10} width={72} />
                            <Tooltip content={<CustomTooltip />} />
                            <Legend />
                            <Line name="Funds Mobilized" type="monotone" dataKey="funds_mobilized" stroke="var(--chart-primary)" strokeWidth={2.4} dot={false} isAnimationActive={false} />
                            <Line name="Net Inflow" type="monotone" dataKey="net_inflow" stroke="var(--chart-secondary)" strokeWidth={2.4} dot={false} isAnimationActive={false} />
                        </LineChart>
                    </ResponsiveContainer>
                </div>
            </Section>

            <Section title="Workbook Preview" subtitle="First rows of the canonical flat AMFI dataset.">
                <DataTable sheet={flatSheet(data)} limit={14} />
            </Section>
        </>
    );
}

function SheetTab({ title, subtitle, sheet }) {
    return (
        <Section title={title} subtitle={subtitle}>
            <DataTable sheet={sheet} />
        </Section>
    );
}

function ArchivesView({ archives, loading, onRefresh }) {
    return (
        <Section title="Finalized Archives" subtitle="Available fiscal years compiled dynamically from the SQLite database layer.">
            <div className="archives-header">
                <button className="btn-primary" onClick={onRefresh} disabled={loading}>
                    <RefreshCw size={16} className={loading ? "spinner" : ""} />
                    Refresh Archives
                </button>
            </div>
            
            {archives.length === 0 ? (
                <div className="empty-state">No archived financial years found.</div>
            ) : (
                <div className="table-scroll">
                    <table className="theory-table" style={{ width: "100%" }}>
                        <thead>
                            <tr>
                                <th>Financial Year</th>
                                <th>Record Count</th>
                                <th>Status</th>
                                <th>Last Modified</th>
                                <th>Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            {archives.map(arch => (
                                <tr key={arch.financial_year}>
                                    <td><strong>FY {arch.financial_year}</strong></td>
                                    <td>{arch.record_count} metrics rows</td>
                                    <td>
                                        <span className={`status-badge ${arch.status.toLowerCase().replace(' ', '-')}`}>
                                            {arch.status}
                                        </span>
                                    </td>
                                    <td>{new Date(arch.last_modified).toLocaleString('en-IN')}</td>
                                    <td>
                                        <a 
                                            href={`${API}/api/download?financial_year=${arch.financial_year}`} 
                                            className="btn-download-archive"
                                            title={`Download FY ${arch.financial_year} Workbook`}
                                        >
                                            <Download size={16} />
                                            <span style={{ marginLeft: "6px" }}>Download Excel</span>
                                        </a>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
        </Section>
    );
}

export default function App() {
    const [activeTab, setActiveTab] = useState('overview');
    const [data, setData] = useState(null);
    const [archives, setArchives] = useState([]);
    const [selectedFY, setSelectedFY] = useState('');
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');
    const [isDarkMode, setIsDarkMode] = useState(true);

    async function loadData(fy) {
        setLoading(true);
        setError('');
        try {
            const url = fy ? `${API}/dashboard-data?financial_year=${fy}` : `${API}/dashboard-data`;
            const res = await fetch(url);
            const payload = await res.json();
            if (!res.ok) throw new Error(payload.detail || 'Unable to load dashboard data.');
            setData(payload);
            if (payload?.financialYear) {
                setSelectedFY(payload.financialYear);
            }
        } catch (err) {
            setError(err.message || 'Unable to load dashboard data.');
        } finally {
            setLoading(false);
        }
    }

    async function loadArchives() {
        try {
            const res = await fetch(`${API}/api/archives`);
            if (res.ok) {
                const list = await res.json();
                setArchives(list);
                if (list.length > 0 && !selectedFY) {
                    const latest = list[0].financial_year;
                    setSelectedFY(latest);
                    loadData(latest);
                }
            }
        } catch (err) {
            console.error("Failed to load archives list:", err);
        }
    }

    async function uploadFile(file) {
        setLoading(true);
        setError('');
        const body = new FormData();
        body.append('file', file);
        try {
            const res = await fetch(`${API}/upload`, { method: 'POST', body });
            const payload = await res.json();
            if (!res.ok) throw new Error(payload.detail || 'Upload failed.');
            setData(payload);
            if (payload.financialYear) {
                setSelectedFY(payload.financialYear);
            }
            loadArchives();
        } catch (err) {
            setError(err.message || 'Upload failed.');
        } finally {
            setLoading(false);
        }
    }

    useEffect(() => {
        loadArchives().then(() => {
            if (!selectedFY) {
                loadData();
            }
        });
    }, []);

    const content = useMemo(() => {
        if (activeTab === 'overview') {
            return <Overview data={data} loading={loading} onUpload={uploadFile} onRefresh={loadData} selectedFY={selectedFY} />;
        }
        if (activeTab === 'sip') {
            return <SheetTab title="SIP Analysis" subtitle="Full historical SIP contribution table." sheet={sipSheet(data)} />;
        }
        if (activeTab === 'flat') {
            return <SheetTab title="Fund Flows (Flat)" subtitle="Scheme-level horizontal master table with all monthly metric blocks." sheet={flatSheet(data)} />;
        }
        if (activeTab === 'form') {
            return <SheetTab title="AMFI Form Layout" subtitle="Regulatory AMFI-style monthly blocks retained in their native structure." sheet={formSheet(data)} />;
        }
        if (activeTab === 'archives') {
            return <ArchivesView archives={archives} loading={loading} onRefresh={loadArchives} />;
        }
        return null;
    }, [activeTab, data, archives, loading, selectedFY]);

    return (
        <div className={`app-layout ${isDarkMode ? 'dark-theme' : 'light-theme'}`}>
            <nav className="sidebar">
                {tabs.map(tab => (
                    <button
                        key={tab.id}
                        className={`sidebar-item ${activeTab === tab.id ? 'active' : ''}`}
                        onClick={() => setActiveTab(tab.id)}
                    >
                        {tab.label}
                    </button>
                ))}
            </nav>
            <main className="main-content">
                <header className="app-header">
                    <div>
                        <h1>AMFI Dashboard</h1>
                        <p>Centralized monthly ingestion and time-series review for mutual fund category data.</p>
                    </div>
                    {archives.length > 0 && (
                        <div className="header-controls">
                            <label htmlFor="fy-select-control" style={{ fontWeight: '700', color: 'var(--text-secondary)', fontSize: '0.9rem' }}>
                                Active FY:
                            </label>
                            <select
                                id="fy-select-control"
                                className="fy-select"
                                value={selectedFY}
                                onChange={e => {
                                    const nextFY = e.target.value;
                                    setSelectedFY(nextFY);
                                    loadData(nextFY);
                                }}
                            >
                                {archives.map(a => (
                                    <option key={a.financial_year} value={a.financial_year}>
                                        FY {a.financial_year}
                                    </option>
                                ))}
                            </select>
                        </div>
                    )}
                </header>
                {error ? (
                    <div className="error-banner" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderLeft: '4px solid #ef4444', backgroundColor: 'rgba(239, 68, 68, 0.1)', color: '#ef4444' }}>
                        <span style={{ fontWeight: '600' }}>Error: {error}</span>
                        <button 
                            onClick={() => setError('')} 
                            style={{ background: 'none', border: 'none', color: 'inherit', cursor: 'pointer', fontWeight: 'bold', fontSize: '1.2rem', marginLeft: '10px' }}
                            title="Dismiss error"
                        >
                            &times;
                        </button>
                    </div>
                ) : null}
                <div className="page-content">{content}</div>
            </main>
            <button
                className="theme-toggle"
                onClick={() => setIsDarkMode(value => !value)}
                title={isDarkMode ? 'Switch to light mode' : 'Switch to dark mode'}
                aria-label={isDarkMode ? 'Switch to light mode' : 'Switch to dark mode'}
            >
                {isDarkMode ? <Sun size={18} /> : <Moon size={18} />}
            </button>
            {selectedFY && (
                <a className="download-fab" href={`${API}/api/download?financial_year=${selectedFY}`} title={`Download FY ${selectedFY} workbook`}>
                    <Download size={19} />
                </a>
            )}
        </div>
    );
}
