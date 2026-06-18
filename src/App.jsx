import React, { useEffect, useId, useMemo, useRef, useState } from 'react';
import {
    Archive,
    Check,
    ChevronDown,
    Download,
    FileSpreadsheet,
    FileUp,
    Layers,
    LayoutDashboard,
    Moon,
    RefreshCw,
    Repeat,
    Search,
    SlidersHorizontal,
    Sun,
    TrendingUp,
} from 'lucide-react';
import {
    Area,
    AreaChart,
    Bar,
    BarChart,
    CartesianGrid,
    Legend,
    Line,
    LineChart,
    Rectangle,
    ResponsiveContainer,
    Tooltip,
    XAxis,
    YAxis,
} from 'recharts';

const API = import.meta.env.VITE_API_URL ||
    ((window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') && window.location.port !== '8000'
        ? 'http://127.0.0.1:8000'
        : window.location.origin);

const tabs = [
    { id: 'overview', label: 'Overview', icon: LayoutDashboard },
    { id: 'categories', label: 'Categories', icon: Layers },
    { id: 'schemes', label: 'Schemes', icon: FileSpreadsheet },
    { id: 'ns', label: 'NS Analysis', icon: TrendingUp },
    { id: 'sip', label: 'SIP', icon: Repeat },
    { id: 'archives', label: 'Archives', icon: Archive },
];

const chartLocale = 'en-IN';
const lineTooltipCursor = { stroke: 'var(--chart-cursor-line)', strokeWidth: 1 };
const categoryFlowSeries = [
    { name: 'Sales', dataKey: 'monthlySales', color: 'var(--chart-primary)' },
    { name: 'Redemption', dataKey: 'monthlyRedemption', color: 'var(--chart-danger)' },
    { name: 'Net Flow', dataKey: 'monthlyNetFlow', color: 'var(--chart-secondary)' },
];
const topAumSeries = [
    { name: 'Latest AUM', dataKey: 'latestAum', color: 'var(--chart-primary)' },
];
const sipAnnualSeries = [
    { name: 'Contribution', dataKey: 'contribution', color: 'var(--chart-primary)' },
];

function nsComparisonSeries(ns) {
    return [
        { name: ns.previousMonth || 'Previous', dataKey: 'previous', color: 'var(--chart-danger)' },
        { name: ns.currentMonth || 'Current', dataKey: 'current', color: 'var(--chart-primary)' },
    ];
}

function renderActiveBar(props) {
    return <Rectangle {...props} stroke="var(--chart-active-stroke)" strokeWidth={2} strokeOpacity={0.85} />;
}

function formatNumber(value, digits = 2) {
    if (typeof value !== 'number' || Number.isNaN(value)) return value ?? '';
    return new Intl.NumberFormat(chartLocale, { maximumFractionDigits: digits }).format(value);
}

function formatPercent(value) {
    if (typeof value !== 'number' || Number.isNaN(value)) return '-';
    return `${formatNumber(value * 100, 1)}%`;
}

function formatCrore(value) {
    if (typeof value !== 'number' || Number.isNaN(value)) return '0';
    return formatNumber(value, value >= 100000 ? 0 : 1);
}

function formatOptionalNumber(value, digits = 2) {
    if (typeof value !== 'number' || Number.isNaN(value)) return '-';
    return formatNumber(value, digits);
}

function safeSeries(data, key) {
    return Array.isArray(data?.[key]) ? data[key] : [];
}

function valueForSort(row, key) {
    const value = row?.[key];
    if (typeof value === 'number') return value;
    if (value == null) return Number.NEGATIVE_INFINITY;
    return String(value).toLowerCase();
}

function sortedRows(rows, sortKey, direction = 'desc') {
    const multiplier = direction === 'asc' ? 1 : -1;
    return [...rows].sort((left, right) => {
        const a = valueForSort(left, sortKey);
        const b = valueForSort(right, sortKey);
        if (a < b) return -1 * multiplier;
        if (a > b) return 1 * multiplier;
        return 0;
    });
}

function monthTickValues(rows, maxTicks = 7) {
    if (!Array.isArray(rows) || rows.length === 0) return undefined;
    const labels = rows.map(row => row.month).filter(Boolean);
    if (labels.length <= maxTicks) return labels;
    const step = Math.ceil((labels.length - 1) / (maxTicks - 1));
    const ticks = labels.filter((_, index) => index % step === 0);
    const latest = labels[labels.length - 1];
    if (!ticks.includes(latest)) ticks.push(latest);
    return ticks;
}

function Section({ title, subtitle, headerAction, children, className = '' }) {
    return (
        <section className={`card ${className}`}>
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

function SortControls({ sortKey, onSortKeyChange, direction, onDirectionChange, options }) {
    return (
        <>
            <GlassSelect
                icon={<SlidersHorizontal size={16} />}
                value={sortKey}
                options={options}
                onChange={onSortKeyChange}
            />
            <button
                type="button"
                className="btn-sm btn-sort-direction"
                onClick={() => onDirectionChange(direction === 'desc' ? 'asc' : 'desc')}
                title={direction === 'desc' ? 'Descending order' : 'Ascending order'}
            >
                {direction === 'desc' ? 'Desc' : 'Asc'}
            </button>
        </>
    );
}

function MetricCard({ label, value, detail, tone = 'neutral' }) {
    return (
        <div className={`metric-card tone-${tone}`}>
            <span>{label}</span>
            <strong>{value}</strong>
            <small style={{ visibility: detail ? 'visible' : 'hidden' }}>{detail || '\u00A0'}</small>
        </div>
    );
}

function EmptyState({ children = 'No data loaded.' }) {
    return <div className="empty-state">{children}</div>;
}

function formatTooltipValue(item) {
    const value = item?.value;
    if (typeof value !== 'number' || Number.isNaN(value)) return value ?? '-';
    const label = `${item?.name || item?.dataKey || ''} ${item?.dataKey || ''}`.toLowerCase();
    const isPercent = label.includes('growth') || label.includes('share') || label.includes('percent') || label.includes('%');
    const number = isPercent ? value * 100 : value;
    const formatted = number.toLocaleString(chartLocale, { maximumFractionDigits: 2 });
    return isPercent ? `${formatted}%` : formatted;
}

function GlassSelect({ icon, value, options, onChange, ariaLabel = 'Select option' }) {
    const [open, setOpen] = useState(false);
    const [activeIndex, setActiveIndex] = useState(0);
    const rootRef = useRef(null);
    const listboxId = useId();
    const selectedIndex = Math.max(0, options.findIndex(option => option.value === value));
    const selected = options[selectedIndex] || options[0];

    useEffect(() => {
        function handleClick(event) {
            if (!rootRef.current?.contains(event.target)) setOpen(false);
        }
        document.addEventListener('mousedown', handleClick);
        return () => document.removeEventListener('mousedown', handleClick);
    }, []);

    useEffect(() => {
        if (open) setActiveIndex(selectedIndex);
    }, [open, selectedIndex]);

    function selectOption(option) {
        if (!option) return;
        onChange(option.value);
        setOpen(false);
    }

    function moveActive(delta) {
        setActiveIndex(index => {
            if (!options.length) return 0;
            return (index + delta + options.length) % options.length;
        });
    }

    function handleKeyDown(event) {
        if (!options.length) return;
        if (event.key === 'ArrowDown') {
            event.preventDefault();
            if (!open) {
                setOpen(true);
                setActiveIndex(selectedIndex);
            } else {
                moveActive(1);
            }
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            if (!open) {
                setOpen(true);
                setActiveIndex(selectedIndex);
            } else {
                moveActive(-1);
            }
        } else if (event.key === 'Home') {
            event.preventDefault();
            setOpen(true);
            setActiveIndex(0);
        } else if (event.key === 'End') {
            event.preventDefault();
            setOpen(true);
            setActiveIndex(options.length - 1);
        } else if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            if (!open) {
                setOpen(true);
                setActiveIndex(selectedIndex);
            } else {
                selectOption(options[activeIndex]);
            }
        } else if (event.key === 'Escape') {
            setOpen(false);
        }
    }

    return (
        <div className="glass-select" ref={rootRef}>
            <button
                type="button"
                className={`glass-select-trigger ${open ? 'open' : ''}`}
                onClick={() => setOpen(value => !value)}
                onKeyDown={handleKeyDown}
                aria-haspopup="listbox"
                aria-expanded={open}
                aria-controls={listboxId}
                aria-label={ariaLabel}
                aria-activedescendant={open ? `${listboxId}-${activeIndex}` : undefined}
            >
                {icon}
                <span>{selected?.label}</span>
                <ChevronDown size={16} className="glass-select-chevron" />
            </button>
            {open && (
                <div className="glass-select-menu" id={listboxId} role="listbox" aria-label={ariaLabel}>
                    {options.map((option, index) => {
                        const active = option.value === value;
                        const highlighted = index === activeIndex;
                        return (
                            <button
                                type="button"
                                key={option.value}
                                id={`${listboxId}-${index}`}
                                className={`glass-select-option ${active ? 'active' : ''} ${highlighted ? 'highlighted' : ''}`}
                                role="option"
                                aria-selected={active}
                                onMouseEnter={() => setActiveIndex(index)}
                                onClick={() => selectOption(option)}
                            >
                                <span>{option.label}</span>
                                {active && <Check size={15} />}
                            </button>
                        );
                    })}
                </div>
            )}
        </div>
    );
}

function tooltipTitleFromPayload(payload, label, titleKey) {
    const data = payload?.[0]?.payload;
    if (titleKey && data?.[titleKey] != null) return data[titleKey];
    return data?.category || data?.schemeName || data?.financialYear || data?.label || data?.month || label;
}

function ChartTooltip({ active, payload, label, titleKey, series }) {
    if (!active || !payload?.length) return null;
    const data = payload[0]?.payload || {};
    const rows = Array.isArray(series) && series.length
        ? series
            .map(item => ({
                ...item,
                value: data[item.dataKey],
                color: item.color,
            }))
            .filter(item => item.value !== undefined && item.value !== null)
        : payload.filter(item => item?.value !== undefined && item?.value !== null);
    if (!rows.length) return null;
    const title = tooltipTitleFromPayload(payload, label, titleKey);
    return (
        <div className="chart-tooltip">
            <strong className="chart-tooltip-title">{title}</strong>
            {rows.map(item => (
                <div className="chart-tooltip-row" key={`${item.dataKey}-${item.name}`}>
                    <span className="chart-tooltip-series">
                        <span
                            className="chart-tooltip-dot"
                            style={{ backgroundColor: item.color || item.stroke || item.fill || '#ffffff' }}
                        />
                        <span>{item.name || item.dataKey}</span>
                    </span>
                    <span className="chart-tooltip-value">{formatTooltipValue(item)}</span>
                </div>
            ))}
        </div>
    );
}

function UploadControl({ loading, onUpload }) {
    const inputRef = useRef(null);
    const [file, setFile] = useState(null);
    const [dragging, setDragging] = useState(false);

    function accept(files) {
        const selected = files?.[0];
        if (selected) setFile(selected);
    }

    async function submitUpload() {
        if (!file) return;
        await onUpload(file);
        if (inputRef.current) inputRef.current.value = '';
        setFile(null);
    }

    return (
        <div className="upload-row compact-upload">
            <div
                className={`drop-zone ${dragging ? 'dragging' : ''}`}
                onDragOver={event => { event.preventDefault(); setDragging(true); }}
                onDragLeave={() => setDragging(false)}
                onDrop={event => { event.preventDefault(); setDragging(false); accept(event.dataTransfer.files); }}
            >
                <FileUp size={20} />
                <div>
                    <strong>{file?.name || 'Drop the AMFI workbook here'}</strong>
                    <span>.xlsx or .xls · or click Browse</span>
                </div>
                <button type="button" className="btn-secondary" onClick={() => inputRef.current?.click()}>Browse</button>
                <input
                    ref={inputRef}
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={event => accept(event.target.files)}
                    hidden
                />
            </div>
            <button className="btn-primary" onClick={submitUpload} disabled={loading || !file}>
                {loading ? <span className="spinner" /> : <FileUp size={18} />}
                Upload
            </button>
        </div>
    );
}

function Overview({ data, loading, isUploading, onUpload, onRefresh, selectedFY, archives }) {
    const summary = data?.summary || {};
    const series = safeSeries(data, 'timeSeries');
    const latest = series[series.length - 1] || {};
    const previous = series[series.length - 2] || {};
    const aumGrowth = previous.net_aum ? (latest.net_aum - previous.net_aum) / previous.net_aum : null;
    const initialLoading = loading && !data;
    const placeholder = initialLoading ? '—' : null;

    return (
        <>
            <Section
                title="Executive Overview"
                subtitle={`FY ${selectedFY || '-'} dashboard`}
                headerAction={
                    <div className="header-controls">
                        {archives.length > 0 && (
                            <GlassSelect
                                value={selectedFY}
                                options={archives.map(item => ({
                                    value: item.financial_year,
                                    label: `FY ${item.financial_year}`,
                                }))}
                                onChange={onRefresh}
                                ariaLabel="Select financial year"
                            />
                        )}
                        <button className="btn-sm" onClick={() => onRefresh(selectedFY)} disabled={loading} title="Refresh data">
                            <RefreshCw size={16} />
                        </button>
                    </div>
                }
            >
                <UploadControl loading={isUploading} onUpload={onUpload} />
                {data?.warnings?.length ? (
                    <div className="warning-list">
                        {data.warnings.map((warning, index) => <span key={index}>{warning}</span>)}
                    </div>
                ) : null}
                <div className="metric-grid">
                    <MetricCard label="Latest Month" value={placeholder ?? (summary.latestMonth || '-')} detail={initialLoading ? '' : 'Compiled period'} />
                    <MetricCard label="Net AUM" value={placeholder ?? formatCrore(summary.latestNetAum)} detail={initialLoading ? '' : `MoM ${formatPercent(aumGrowth)}`} tone={initialLoading ? 'neutral' : (aumGrowth >= 0 ? 'good' : 'soft')} />
                    <MetricCard label="Funds Mobilized" value={placeholder ?? formatCrore(summary.latestFundsMobilized)} detail={initialLoading ? '' : 'Latest month'} />
                    <MetricCard label="Net Inflow" value={placeholder ?? formatCrore(summary.latestNetInflow)} detail={initialLoading ? '' : 'Sales less redemption'} tone={initialLoading ? 'neutral' : (summary.latestNetInflow >= 0 ? 'good' : 'soft')} />
                </div>
            </Section>

            <div className="dashboard-grid two-col">
                <Section title="AUM Trend">
                    {series.length ? (
                        <div className="chart-frame">
                            <ResponsiveContainer width="100%" height="100%">
                                <AreaChart data={series} margin={{ top: 10, right: 34, left: 4, bottom: 0 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                    <XAxis
                                        dataKey="month"
                                        stroke="var(--chart-axis)"
                                        tickMargin={10}
                                        interval={0}
                                        ticks={monthTickValues(series, 6)}
                                        padding={{ left: 12, right: 36 }}
                                        height={44}
                                    />
                                    <YAxis stroke="var(--chart-axis)" tickMargin={10} width={78} tickFormatter={value => formatCrore(value)} />
                                    <Tooltip content={<ChartTooltip />} cursor={lineTooltipCursor} isAnimationActive={false} />
                                    <Area name="Net AUM" type="monotone" dataKey="net_aum" stroke="var(--chart-primary)" fill="var(--chart-fill)" strokeWidth={2.2} isAnimationActive={false} />
                                </AreaChart>
                            </ResponsiveContainer>
                        </div>
                    ) : <EmptyState>{initialLoading ? 'Loading dashboard…' : 'No data loaded.'}</EmptyState>}
                </Section>

                <Section title="Monthly Flows">
                    {series.length ? (
                        <div className="chart-frame">
                            <ResponsiveContainer width="100%" height="100%">
                                <LineChart data={series} margin={{ top: 10, right: 34, left: 4, bottom: 0 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                    <XAxis
                                        dataKey="month"
                                        stroke="var(--chart-axis)"
                                        tickMargin={10}
                                        interval={0}
                                        ticks={monthTickValues(series, 6)}
                                        padding={{ left: 12, right: 36 }}
                                        height={44}
                                    />
                                    <YAxis stroke="var(--chart-axis)" tickMargin={10} width={78} tickFormatter={value => formatCrore(value)} />
                                    <Tooltip content={<ChartTooltip />} cursor={lineTooltipCursor} isAnimationActive={false} />
                                    <Legend />
                                    <Line name="Sales" type="monotone" dataKey="funds_mobilized" stroke="var(--chart-primary)" strokeWidth={2.2} dot={false} isAnimationActive={false} />
                                    <Line name="Redemption" type="monotone" dataKey="redemption" stroke="var(--chart-danger)" strokeWidth={2.2} dot={false} isAnimationActive={false} />
                                    <Line name="Net Flow" type="monotone" dataKey="net_inflow" stroke="var(--chart-secondary)" strokeWidth={2.2} dot={false} isAnimationActive={false} />
                                </LineChart>
                            </ResponsiveContainer>
                        </div>
                    ) : <EmptyState>{initialLoading ? 'Loading dashboard…' : 'No data loaded.'}</EmptyState>}
                </Section>
            </div>
        </>
    );
}

function CategoriesView({ data }) {
    const categories = safeSeries(data, 'categorySummary');
    const [sortKey, setSortKey] = useState('latestAum');
    const [sortDirection, setSortDirection] = useState('desc');
    const sortedCategories = useMemo(
        () => sortedRows(categories, sortKey, sortDirection),
        [categories, sortKey, sortDirection],
    );
    const totalAum = categories.reduce((sum, row) => sum + (row.latestAum || 0), 0);
    const sortOptions = [
        { value: 'latestAum', label: 'Latest AUM' },
        { value: 'monthlySales', label: 'Sales' },
        { value: 'monthlyNetFlow', label: 'Net Flow' },
        { value: 'fytdSales', label: 'FYTD Sales' },
        { value: 'aumGrowth', label: 'AUM Growth' },
    ];

    return (
        <>
            <Section
                title="Category Allocation"
                subtitle="AUM share and latest monthly flow by fund type"
                headerAction={
                    <div className="table-controls">
                        <SortControls
                            sortKey={sortKey}
                            onSortKeyChange={setSortKey}
                            direction={sortDirection}
                            onDirectionChange={setSortDirection}
                            options={sortOptions}
                        />
                    </div>
                }
            >
                {sortedCategories.length ? (
                    <div className="category-layout">
                        <div className="allocation-list">
                            {sortedCategories.map(row => (
                                <div className="allocation-row" key={row.category}>
                                    <div>
                                        <strong>{row.category}</strong>
                                        <span>{formatCrore(row.latestAum)} Cr AUM</span>
                                    </div>
                                    <div className="allocation-bar" aria-label={`${row.category} AUM share`}>
                                        <span style={{ width: `${Math.max(row.aumShare * 100, 2)}%` }} />
                                    </div>
                                    <em>{formatPercent(row.aumShare)}</em>
                                </div>
                            ))}
                        </div>
                        <div className="category-total">
                            <span>Total AUM</span>
                            <strong>{formatCrore(totalAum)} Cr</strong>
                            <small>{sortedCategories.length} categories</small>
                        </div>
                    </div>
                ) : <EmptyState />}
            </Section>

            <Section title="Sales, Redemption, Net Flow">
                {sortedCategories.length ? (
                    <div className="chart-frame tall">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={sortedCategories} margin={{ top: 10, right: 22, left: 6, bottom: 24 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                <XAxis dataKey="category" stroke="var(--chart-axis)" interval={0} tickMargin={10} height={70} />
                                <YAxis stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} width={78} />
                                <Tooltip
                                    content={<ChartTooltip titleKey="category" series={categoryFlowSeries} />}
                                    cursor={false}
                                    isAnimationActive={false}
                                />
                                <Legend />
                                <Bar name="Sales" dataKey="monthlySales" fill="var(--chart-primary)" activeBar={renderActiveBar} isAnimationActive={false} />
                                <Bar name="Redemption" dataKey="monthlyRedemption" fill="var(--chart-danger)" activeBar={renderActiveBar} isAnimationActive={false} />
                                <Bar name="Net Flow" dataKey="monthlyNetFlow" fill="var(--chart-secondary)" activeBar={renderActiveBar} isAnimationActive={false} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                ) : <EmptyState />}
            </Section>

            <Section title="Category Summary">
                <SummaryTable
                    columns={[
                        ['category', 'Category'],
                        ['schemeCount', 'Schemes'],
                        ['latestAum', 'Latest AUM'],
                        ['monthlySales', 'Sales'],
                        ['monthlyRedemption', 'Redemption'],
                        ['monthlyNetFlow', 'Net Flow'],
                        ['fytdSales', 'FYTD Sales'],
                        ['aumGrowth', 'AUM Growth'],
                    ]}
                    rows={sortedCategories}
                />
            </Section>
        </>
    );
}

function SchemesView({ data }) {
    const schemes = safeSeries(data, 'schemeSummary');
    const categories = useMemo(() => [...new Set(schemes.map(row => row.category))], [schemes]);
    const categoryOptions = useMemo(() => [
        { value: 'All', label: 'All categories' },
        ...categories.map(item => ({ value: item, label: item })),
    ], [categories]);
    const [query, setQuery] = useState('');
    const [category, setCategory] = useState('All');
    const [sortKey, setSortKey] = useState('latestAum');
    const [sortDirection, setSortDirection] = useState('desc');
    const sortOptions = [
        { value: 'latestAum', label: 'Latest AUM' },
        { value: 'monthlySales', label: 'Sales' },
        { value: 'monthlyRedemption', label: 'Redemption' },
        { value: 'monthlyNetFlow', label: 'Net Flow' },
        { value: 'fytdSales', label: 'FYTD Sales' },
        { value: 'aumGrowth', label: 'AUM Growth' },
    ];

    const filtered = useMemo(() => {
        const q = query.trim().toLowerCase();
        return schemes.filter(row => {
            const matchesQuery = !q || `${row.schemeName} ${row.fundType} ${row.category}`.toLowerCase().includes(q);
            const matchesCategory = category === 'All' || row.category === category;
            return matchesQuery && matchesCategory;
        });
    }, [schemes, query, category]);

    const sortedFiltered = useMemo(
        () => sortedRows(filtered, sortKey, sortDirection),
        [filtered, sortKey, sortDirection],
    );
    const topSchemes = sortedFiltered.slice(0, 12);

    return (
        <>
            <Section
                title="Scheme Drill-Down"
                subtitle={`${filtered.length} of ${schemes.length} schemes`}
                headerAction={
                    <div className="table-controls">
                        <label className="search-box">
                            <Search size={16} />
                            <input value={query} onChange={event => setQuery(event.target.value)} placeholder="Search schemes" />
                        </label>
                        <GlassSelect
                            icon={<SlidersHorizontal size={16} />}
                            value={category}
                            options={categoryOptions}
                            onChange={setCategory}
                        />
                        <SortControls
                            sortKey={sortKey}
                            onSortKeyChange={setSortKey}
                            direction={sortDirection}
                            onDirectionChange={setSortDirection}
                            options={sortOptions}
                        />
                    </div>
                }
            >
                <SummaryTable
                    columns={[
                        ['schemeName', 'Scheme'],
                        ['category', 'Category'],
                        ['latestAum', 'Latest AUM'],
                        ['monthlySales', 'Sales'],
                        ['monthlyRedemption', 'Redemption'],
                        ['monthlyNetFlow', 'Net Flow'],
                        ['averageAum', 'Avg AUM'],
                        ['fytdSales', 'FYTD Sales'],
                        ['aumGrowth', 'AUM Growth'],
                    ]}
                    rows={sortedFiltered}
                />
            </Section>

            <Section title="Top AUM Schemes">
                {topSchemes.length ? (
                    <div className="chart-frame tall">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={topSchemes} layout="vertical" margin={{ top: 8, right: 28, left: 132, bottom: 8 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                <XAxis type="number" stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} />
                                <YAxis type="category" dataKey="schemeName" stroke="var(--chart-axis)" width={132} />
                                <Tooltip
                                    content={<ChartTooltip titleKey="schemeName" series={topAumSeries} />}
                                    cursor={false}
                                    isAnimationActive={false}
                                />
                                <Bar name="Latest AUM" dataKey="latestAum" fill="var(--chart-primary)" activeBar={renderActiveBar} isAnimationActive={false} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                ) : <EmptyState />}
            </Section>
        </>
    );
}

function SipView({ data }) {
    const sip = data?.sipSummary || {};
    const monthly = Array.isArray(sip.monthlySeries) ? sip.monthlySeries : [];
    const annual = Array.isArray(sip.annualContributions) ? sip.annualContributions : [];
    const sipStatsRows = Array.isArray(sip.sipStatsRows) ? sip.sipStatsRows : [];
    const [sipStatsSortKey, setSipStatsSortKey] = useState('templateOrder');
    const [sipStatsSortDirection, setSipStatsSortDirection] = useState('desc');
    const sipStatsSortOptions = [
        { value: 'templateOrder', label: 'Default' },
        { value: 'label', label: 'Month / Period' },
        { value: 'newRegistrations', label: 'New SIPs' },
        { value: 'discontinued', label: 'Discontinued' },
        { value: 'contributingAccounts', label: 'Contributing SIPs' },
    ];
    const sortedSipStatsRows = useMemo(
        () => sipStatsSortKey === 'templateOrder'
            ? sipStatsRows
            : sortedRows(sipStatsRows, sipStatsSortKey, sipStatsSortDirection),
        [sipStatsRows, sipStatsSortKey, sipStatsSortDirection],
    );
    const sipAumDetail = sip.latestAumMonth ? `Latest available: ${sip.latestAumMonth}` : 'Latest available';
    const sipAccountsDetail = sip.latestOutstandingAccountsMonth ? `Latest available: ${sip.latestOutstandingAccountsMonth}` : 'Latest available';

    return (
        <>
            <Section
                title="SIP Snapshot"
                subtitle={`Latest: ${sip.latestMonth || '-'}`}
                headerAction={
                    <div className="table-controls">
                        <SortControls
                            sortKey={sipStatsSortKey}
                            onSortKeyChange={setSipStatsSortKey}
                            direction={sipStatsSortDirection}
                            onDirectionChange={setSipStatsSortDirection}
                            options={sipStatsSortOptions}
                        />
                    </div>
                }
            >
                <div className="sip-snapshot-layout">
                    <div className="metric-grid sip-core-grid">
                        <MetricCard label="Contribution" value={`${formatCrore(sip.latestContribution || 0)} Cr`} detail={`MoM ${formatPercent(sip.contributionGrowth)}`} />
                        <MetricCard label="SIP AUM" value={`${formatCrore(sip.latestAum || 0)} Cr`} detail={sipAumDetail} />
                        <MetricCard label="Outstanding Accounts" value={`${formatNumber(sip.latestOutstandingAccounts || 0, 2)} Lakh`} detail={sipAccountsDetail} />
                        <MetricCard label="Months Tracked" value={String(monthly.length)} />
                    </div>
                    <SipStatsTable rows={sortedSipStatsRows} />
                </div>
            </Section>

            <div className="dashboard-grid two-col">
                <Section title="SIP Contribution Trend">
                    {monthly.length ? (
                        <div className="chart-frame">
                            <ResponsiveContainer width="100%" height="100%">
                                <LineChart data={monthly} margin={{ top: 10, right: 22, left: 4, bottom: 0 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                    <XAxis
                                        dataKey="month"
                                        stroke="var(--chart-axis)"
                                        tickMargin={12}
                                        interval={0}
                                        ticks={monthTickValues(monthly, 7)}
                                        height={64}
                                        angle={-35}
                                        textAnchor="end"
                                    />
                                    <YAxis stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} width={78} />
                                    <Tooltip content={<ChartTooltip />} cursor={lineTooltipCursor} isAnimationActive={false} />
                                    <Line name="Contribution" type="monotone" dataKey="contribution" stroke="var(--chart-primary)" strokeWidth={2.2} dot={false} isAnimationActive={false} />
                                </LineChart>
                            </ResponsiveContainer>
                        </div>
                    ) : <EmptyState />}
                </Section>

                <Section title="SIP AUM">
                    {monthly.length ? (
                        <div className="chart-frame">
                            <ResponsiveContainer width="100%" height="100%">
                                <AreaChart data={monthly} margin={{ top: 10, right: 22, left: 4, bottom: 0 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                    <XAxis
                                        dataKey="month"
                                        stroke="var(--chart-axis)"
                                        tickMargin={12}
                                        interval={0}
                                        ticks={monthTickValues(monthly, 7)}
                                        height={64}
                                        angle={-35}
                                        textAnchor="end"
                                    />
                                    <YAxis stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} width={78} />
                                    <Tooltip content={<ChartTooltip />} cursor={lineTooltipCursor} isAnimationActive={false} />
                                    <Area name="SIP AUM" type="monotone" dataKey="aum" stroke="var(--chart-secondary)" fill="var(--chart-fill)" strokeWidth={2.2} isAnimationActive={false} />
                                </AreaChart>
                            </ResponsiveContainer>
                        </div>
                    ) : <EmptyState />}
                </Section>
            </div>

            <Section title="Annual SIP Contributions">
                {annual.length ? (
                    <div className="chart-frame tall">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={annual} margin={{ top: 10, right: 22, left: 4, bottom: 24 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                <XAxis dataKey="financialYear" stroke="var(--chart-axis)" tickMargin={10} interval={0} />
                                <YAxis stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} width={78} />
                                <Tooltip
                                    content={<ChartTooltip titleKey="financialYear" series={sipAnnualSeries} />}
                                    cursor={false}
                                    isAnimationActive={false}
                                />
                                <Bar name="Contribution" dataKey="contribution" fill="var(--chart-primary)" activeBar={renderActiveBar} isAnimationActive={false} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                ) : <EmptyState />}
            </Section>
        </>
    );
}

function SipStatsTable({ rows }) {
    if (!rows?.length) return <EmptyState>No SIP count rows available.</EmptyState>;
    return (
        <div className="table-scroll compact sip-stats-table-wrap">
            <table className="theory-table sip-stats-table">
                <thead>
                    <tr>
                        <th>Month / Period</th>
                        <th>No. of New SIPs registered</th>
                        <th>No. of SIPs discontinued*/tenure completed</th>
                        <th>No. of Contributing SIP accounts</th>
                    </tr>
                </thead>
                <tbody>
                    {rows.map((row, index) => (
                        <tr key={`${row.label}-${index}`} className={row.isSummary ? 'summary-row' : undefined}>
                            <td>{row.label}</td>
                            <td className="numeric-cell">{formatOptionalNumber(row.newRegistrations, 2)}</td>
                            <td className="numeric-cell">{formatOptionalNumber(row.discontinued, 2)}</td>
                            <td className="numeric-cell">{formatOptionalNumber(row.contributingAccounts, 2)}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}

function NsAnalysisView({ data }) {
    const ns = data?.nsAnalysis || {};
    const headlineRows = Array.isArray(ns.headlineRows) ? ns.headlineRows : [];
    const equityRows = Array.isArray(ns.equityRows) ? ns.equityRows : [];
    const hybridRows = Array.isArray(ns.hybridRows) ? ns.hybridRows : [];
    const total = ns.total || {};
    const period = `${ns.previousMonth || '-'} to ${ns.currentMonth || '-'}`;
    const comparisonRows = [total, ...headlineRows].filter(row => row?.label);
    const [sortKey, setSortKey] = useState('current');
    const [sortDirection, setSortDirection] = useState('desc');
    const schemeRows = [
        ...equityRows,
        ...(ns.equityTotal ? [ns.equityTotal] : []),
        ...hybridRows,
        ...(ns.hybridTotal ? [ns.hybridTotal] : []),
    ];
    const sortedSchemeRows = useMemo(
        () => sortedRows(schemeRows, sortKey, sortDirection),
        [schemeRows, sortKey, sortDirection],
    );
    const sortOptions = [
        { value: 'current', label: ns.currentMonth || 'Current' },
        { value: 'previous', label: ns.previousMonth || 'Previous' },
        { value: 'growth', label: 'Growth %' },
        { value: 'label', label: 'Segment' },
    ];

    const nsColumns = [
        ['label', 'Segment'],
        ['previous', ns.previousMonth || 'Previous'],
        ['current', ns.currentMonth || 'Current'],
        ['growth', 'Growth %'],
    ];

    return (
        <>
            <Section title="NS Analysis" subtitle={`Net sales comparison: ${period}`}>
                <div className="metric-grid">
                    <MetricCard label="Total Net Sales" value={`${formatCrore(total.current || 0)} Cr`} detail={`${ns.currentMonth || 'Current'} | Growth ${formatPercent(total.growth)}`} />
                    {headlineRows.map(row => (
                        <MetricCard
                            key={row.label}
                            label={row.label}
                            value={`${formatCrore(row.current || 0)} Cr`}
                            detail={`${ns.previousMonth || 'Previous'}: ${formatCrore(row.previous || 0)} Cr`}
                            tone={(row.current || 0) >= 0 ? 'good' : 'soft'}
                        />
                    ))}
                </div>
            </Section>

            <Section title="Previous vs Current">
                {comparisonRows.length ? (
                    <div className="chart-frame tall">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={comparisonRows} margin={{ top: 10, right: 22, left: 6, bottom: 24 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="var(--chart-grid)" />
                                <XAxis dataKey="label" stroke="var(--chart-axis)" tickMargin={10} interval={0} />
                                <YAxis stroke="var(--chart-axis)" tickFormatter={value => formatCrore(value)} width={78} />
                                <Tooltip
                                    content={<ChartTooltip titleKey="label" series={nsComparisonSeries(ns)} />}
                                    cursor={false}
                                    isAnimationActive={false}
                                />
                                <Legend />
                                <Bar name={ns.previousMonth || 'Previous'} dataKey="previous" fill="var(--chart-danger)" activeBar={renderActiveBar} isAnimationActive={false} />
                                <Bar name={ns.currentMonth || 'Current'} dataKey="current" fill="var(--chart-primary)" activeBar={renderActiveBar} isAnimationActive={false} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                ) : <EmptyState />}
            </Section>

            <Section
                title="NS Breakdown"
                headerAction={
                    <div className="table-controls">
                        <SortControls
                            sortKey={sortKey}
                            onSortKeyChange={setSortKey}
                            direction={sortDirection}
                            onDirectionChange={setSortDirection}
                            options={sortOptions}
                        />
                    </div>
                }
            >
                <SummaryTable columns={nsColumns} rows={sortedSchemeRows} />
            </Section>
        </>
    );
}

function SummaryTable({ columns, rows }) {
    if (!rows?.length) return <EmptyState>No rows available.</EmptyState>;
    return (
        <div className="table-scroll compact">
            <table className="theory-table">
                <thead>
                    <tr>
                        {columns.map(([, label]) => <th key={label}>{label}</th>)}
                    </tr>
                </thead>
                <tbody>
                    {rows.map((row, rowIndex) => (
                        <tr key={row.schemeKey || row.category || rowIndex}>
                            {columns.map(([key]) => {
                                const value = row[key];
                                const isPercent = key.toLowerCase().includes('growth') || key.toLowerCase().includes('share');
                                const isNumeric = typeof value === 'number';
                                return (
                                    <td key={key} className={isNumeric ? 'numeric-cell' : undefined}>
                                        {isPercent ? formatPercent(value) : isNumeric ? formatNumber(value) : value}
                                    </td>
                                );
                            })}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}

function ArchivesView({ archives, loading, selectedFY }) {
    return (
        <Section title="Archives" subtitle="Download generated workbooks by financial year">
            {archives.length === 0 ? (
                <EmptyState>No archived financial years found.</EmptyState>
            ) : (
                <div className="table-scroll compact">
                    <table className="theory-table archive-table">
                        <thead>
                            <tr>
                                <th>Financial Year</th>
                                <th>Records</th>
                                <th>Status</th>
                                <th>Last Modified</th>
                                <th>Download</th>
                            </tr>
                        </thead>
                        <tbody>
                            {archives.map(item => (
                                <tr key={item.financial_year}>
                                    <td>FY {item.financial_year}</td>
                                    <td className="numeric-cell">{formatNumber(item.record_count, 0)}</td>
                                    <td>
                                        <span className={`status-badge ${String(item.status || '').toLowerCase().replace(/\s+/g, '-')}`}>
                                            {item.status || '-'}
                                        </span>
                                    </td>
                                    <td>{item.last_modified ? new Date(item.last_modified).toLocaleString('en-IN') : '-'}</td>
                                    <td>
                                        <a
                                            className="btn-download-archive"
                                            href={`${API}/api/download?financial_year=${encodeURIComponent(item.financial_year)}`}
                                            title={`Download FY ${item.financial_year} workbook`}
                                        >
                                            <Download size={15} />
                                            Excel
                                        </a>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
            {loading && <div className="archive-actions"><span className="muted-text">Refreshing...</span></div>}
        </Section>
    );
}

export default function App() {
    const [activeTab, setActiveTab] = useState('overview');
    const [data, setData] = useState(null);
    const [archives, setArchives] = useState([]);
    const [selectedFY, setSelectedFY] = useState('');
    const [loading, setLoading] = useState(false);
    const [isUploading, setIsUploading] = useState(false);
    const [error, setError] = useState('');
    const [isDarkMode, setIsDarkMode] = useState(true);

    useEffect(() => {
        const theme = isDarkMode ? 'dark' : 'light';
        document.documentElement.dataset.theme = theme;
        document.documentElement.style.colorScheme = theme;
    }, [isDarkMode]);

    async function loadData(fy) {
        setLoading(true);
        setError('');
        try {
            const url = fy ? `${API}/dashboard-data?financial_year=${fy}` : `${API}/dashboard-data`;
            const res = await fetch(url);
            const payload = await res.json();
            if (!res.ok) throw new Error(payload.detail || 'Unable to load dashboard data.');
            setData(payload);
            if (payload?.financialYear) setSelectedFY(payload.financialYear);
        } catch (err) {
            setError(err.message || 'Unable to load dashboard data.');
        } finally {
            setLoading(false);
        }
    }

    async function loadArchives() {
        try {
            const res = await fetch(`${API}/api/archives`);
            if (!res.ok) return;
            const list = await res.json();
            setArchives(list);
        } catch (err) {
            console.error('Failed to load archives list:', err);
        }
    }

    async function uploadFile(file) {
        setIsUploading(true);
        setError('');
        const body = new FormData();
        body.append('file', file);
        try {
            const res = await fetch(`${API}/upload`, { method: 'POST', body });
            const payload = await res.json();
            if (!res.ok) throw new Error(payload.detail || 'Upload failed.');
            setData(payload);
            if (payload.financialYear) setSelectedFY(payload.financialYear);
            loadArchives();
        } catch (err) {
            setError(err.message || 'Upload failed.');
        } finally {
            setIsUploading(false);
        }
    }

    useEffect(() => {
        // Critical dashboard data fires immediately; archives load in parallel so
        // the FY dropdown never blocks the metric cards or charts.
        loadData();
        loadArchives();
    }, []);

    const content = useMemo(() => {
        if (activeTab === 'overview') {
            return <Overview data={data} loading={loading} isUploading={isUploading} onUpload={uploadFile} onRefresh={loadData} selectedFY={selectedFY} archives={archives} />;
        }
        if (activeTab === 'categories') return <CategoriesView data={data} />;
        if (activeTab === 'schemes') return <SchemesView data={data} />;
        if (activeTab === 'ns') return <NsAnalysisView data={data} />;
        if (activeTab === 'sip') return <SipView data={data} />;
        if (activeTab === 'archives') return <ArchivesView archives={archives} loading={loading} selectedFY={selectedFY} />;
        return null;
    }, [activeTab, data, archives, loading, isUploading, selectedFY]);

    return (
        <div className={`app-layout ${isDarkMode ? 'dark-theme' : 'light-theme'}`}>
            <nav className="sidebar">
                {tabs.map(tab => (
                    <button
                        key={tab.id}
                        className={`sidebar-item ${activeTab === tab.id ? 'active' : ''}`}
                        onClick={() => setActiveTab(tab.id)}
                    >
                        <tab.icon size={17} /><span>{tab.label}</span>
                    </button>
                ))}
            </nav>
            <main className="main-content">
                <header className="app-header">
                    <div>
                        <h1>AMFI Dashboard</h1>
                        <p>Mutual fund category flows, AUM movement, scheme drill-down, and SIP trend review.</p>
                    </div>
                </header>
                {error ? (
                    <div className="error-banner">
                        <span>Error: {error}</span>
                        <button onClick={() => setError('')} title="Dismiss error">&times;</button>
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
