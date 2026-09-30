import streamlit as st
import pandas as pd
import re
import time
from datetime import datetime
from threat_hunter import fetch_url, parse_article, generate_excel_report, extract_article_links, parse_cisa_kev, group_findings

# Default Cybersecurity Sources from analyst work list
DEFAULT_SOURCES = {
    "Ars Technica": "https://arstechnica.com/",
    "Krebs on Security": "https://krebsonsecurity.com/",
    "SC Magazine": "https://www.scmagazine.com/",
    "Security Week": "https://www.securityweek.com/",
    "Bleeping Computer": "https://www.bleepingcomputer.com/",
    "The Hacker News": "https://thehackernews.com",
    "Cyber News": "https://cybernews.com",
    "SC Media - Cyber Magazine": "https://cybermagazine.com",
    "CISA Topics & Advisories": "https://www.cisa.gov/topics/cyber-threats-and-advisories",
    "CISA Known Exploited Vulnerabilities (KEV) Catalog": "https://www.cisa.gov/known-exploited-vulnerabilities-catalog",
    "Security Gladiators": "https://securitygladiators.com/cybersecurity/",
    "IT Security Guru": "https://www.ITSecurityGuru.org",
    "Graham Cluley": "https://grahamcluley.com/",
    "Naked Security (Sophos)": "https://news.sophos.com/en-us/category/serious-security/",
    "The Last Watchdog": "https://www.lastwatchdog.com/",
    "State of Security (Tripwire)": "https://www.tripwire.com/state-of-security",
    "CSO Online UK": "https://www.csoonline.com/uk/",
    "InfoSecurity Magazine": "https://www.infosecurity-magazine.com/",
    "Recorded Future": "https://www.recordedfuture.com/recorded-future-news",
    "National Vulnerability Database": "https://nvd.nist.gov/vuln/search",
    "National Mortgage News": "https://www.nationalmortgagenews.com/",
    "SentinelOne Vulnerability Database": "https://www.sentinelone.com/vulnerability-database/",
    "SOC Prime Blog": "https://socprime.com/blog/",
    "Wordfence Blog": "https://www.wordfence.com/blog/",
    "NCSC Alerts": "https://www.ncsc.govt.nz/alerts/",
    "Safe Computing Security Alerts (U-M)": "https://safecomputing.umich.edu/security-alerts"
}

# Set page config
st.set_page_config(
    page_title="Ck's Threat Hunter application",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern Cyber Security UI Styling (Premium Royal Blue Light Mode Theme)
st.markdown("""
<style>
    /* Main Layout and Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Style main Streamlit container background (Clean white/gray) */
    .stApp {
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    
    /* Style sidebar background (Light slate) */
    section[data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
        border-right: 1px solid #cbd5e1 !important;
    }
    
    /* Code / Monospace styling */
    code, pre {
        font-family: 'JetBrains Mono', monospace !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
    }
    
    /* Title Styling */
    .title-container {
        padding: 1.5rem 2rem;
        background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%);
        border-radius: 12px;
        border-left: 5px solid #2563eb;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        border-top: 1px solid #e2e8f0;
        border-right: 1px solid #e2e8f0;
        border-bottom: 1px solid #e2e8f0;
    }
    .title-text {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .subtitle-text {
        font-size: 1.05rem;
        color: #475569;
        margin-top: 0.5rem;
    }
    
    /* Sidebar styling */
    .sidebar-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: #0f172a;
        margin-bottom: 1rem;
        border-bottom: 1px solid #cbd5e1;
        padding-bottom: 0.5rem;
    }
    
    /* Custom buttons and forms */
    .stButton>button {
        background-color: #2563eb !important;
        color: white !important;
        border: none !important;
        padding: 0.5rem 1.5rem !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        width: 100%;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.15);
    }
    .stButton>button:hover {
        background-color: #1d4ed8 !important;
        transform: translateY(-1px);
        box-shadow: 0 6px 15px rgba(37, 99, 235, 0.3);
    }
    
    /* Style inputs */
    div[data-baseweb="textarea"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }
    textarea {
        color: #0f172a !important;
    }
    div[data-baseweb="base-input"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }
    input {
        color: #0f172a !important;
    }
    
    /* Info/Status Cards */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #2563eb;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Log block style */
    .log-container {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 1rem;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.85rem;
        color: #334155;
        max-height: 250px;
        overflow-y: auto;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Application Header
st.markdown("""
<div class="title-container">
    <div class="title-text">🛡️ Ck's Threat Hunter application</div>
    <div class="subtitle-text">Automate extraction of CVEs and vulnerability metadata from news articles, threat advisories, and blog posts.</div>
</div>
""", unsafe_allow_html=True)

# Sidebar Configuration & Help
with st.sidebar:
    st.markdown('<div class="sidebar-header">Threat Intelligence Feeds</div>', unsafe_allow_html=True)
    st.write("Click on any of the popular security portals below to copy the link and add it to your list:")
    
    portal_links = {
        "CISA KEV Catalog": "https://www.cisa.gov/known-exploited-vulnerabilities-catalog",
        "CISA Alerts": "https://www.cisa.gov/news-events/cybersecurity-advisories",
        "BleepingComputer": "https://www.bleepingcomputer.com/news/security/",
        "The Hacker News": "https://thehackernews.com/",
        "NIST NVD Updates": "https://nvd.nist.gov/vuln",
        "Palo Alto Unit 42": "https://unit42.paloaltonetworks.com/",
        "Dark Reading": "https://www.darkreading.com/"
    }

    for name, link in portal_links.items():
        st.code(link, language="text")

    st.markdown("---")
    st.markdown('<div class="sidebar-header">App Capabilities</div>', unsafe_allow_html=True)
    st.info(
        "🔍 **Regex Scan**: Finds exact CVE identifiers in page content.\n\n"
        "💡 **Smart Context**: Pinpoints the exact paragraph containing the CVE to write a description.\n\n"
        "📊 **Styled Output**: Auto-formats Excel tables with cell borders, word wrapping, and clear sizes."
    )

# Main Dashboard layout split into Input and Execution
col1, col2 = st.columns([1, 1])

# Keep track of list of URLs
urls_to_scan = []

with col1:
    st.subheader("1. Enter Sources")
    
    # Checkbox to include default feeds
    include_defaults = st.checkbox(f"Include Default Cybersecurity Sources ({len(DEFAULT_SOURCES)} feeds)", value=True)
    
    selected_default_urls = []
    if include_defaults:
        selected_names = st.multiselect(
            "Choose default feeds to scan:",
            options=list(DEFAULT_SOURCES.keys()),
            default=list(DEFAULT_SOURCES.keys()),
            help="Uncheck any feeds you do not want to scan in this run."
        )
        selected_default_urls = [DEFAULT_SOURCES[name] for name in selected_names]
        
    st.write("") # spacing
    
    # Text input method for custom URLs
    url_input_text = st.text_area(
        "Paste Custom URLs (one per line):",
        placeholder="https://example-security-site.com/vulnerability-report\nhttps://cisa.gov/news-events/alerts/...",
        height=150
    )
    
    # File upload method for custom URLs
    uploaded_file = st.file_uploader(
        "Or upload a file containing custom URLs (.txt or .csv)",
        type=["txt", "csv"]
    )
    
    # Parse inputs
    parsed_urls = []
    
    # Add selected defaults
    parsed_urls.extend(selected_default_urls)
    
    # Add pasted URLs
    if url_input_text.strip():
        parsed_urls.extend([line.strip() for line in url_input_text.split("\n") if line.strip()])
        
    # Add uploaded file URLs
    if uploaded_file is not None:
        if uploaded_file.name.endswith(".txt"):
            content = uploaded_file.read().decode("utf-8")
            parsed_urls.extend([line.strip() for line in content.split("\n") if line.strip()])
        elif uploaded_file.name.endswith(".csv"):
            df_csv = pd.read_csv(uploaded_file)
            # Search for columns that might contain URLs
            url_col = None
            for col in df_csv.columns:
                if 'url' in col.lower() or 'link' in col.lower():
                    url_col = col
                    break
            
            if url_col:
                parsed_urls.extend(df_csv[url_col].dropna().astype(str).str.strip().tolist())
            else:
                # Fallback: take the first column
                parsed_urls.extend(df_csv.iloc[:, 0].dropna().astype(str).str.strip().tolist())
                
    # Deduplicate and validate basic pattern
    final_urls = []
    for u in parsed_urls:
        if u and not u.startswith('#') and (u.startswith('http') or '.' in u):
            final_urls.append(u)
    final_urls = list(dict.fromkeys(final_urls)) # deduplicate keeping order
    
    st.write(f"**Total unique URLs loaded:** `{len(final_urls)}`")

with col2:
    st.subheader("2. Start Intelligence Collection")
    
    # Start button
    start_button = st.button("🚀 Start Threat Hunt", disabled=len(final_urls) == 0)
    
    if start_button:
        st.session_state['running'] = True
        st.session_state['urls'] = final_urls
        st.session_state['results'] = []
        st.session_state['current_index'] = 0

if 'running' in st.session_state and st.session_state['running']:
    urls = st.session_state['urls']
    results = st.session_state['results']
    idx = st.session_state['current_index']
    
    # Progress indicators
    progress_bar = st.progress(0.0)
    status_text = st.empty()
    
    # Logging text area
    log_header = st.write("**Scanning logs:**")
    log_box = st.empty()
    logs = []
    
    logs.append(f"[*] Starting scan for {len(urls)} URLs...")
    log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
    
    # Start threat hunting iteration
    results_list = []
    
    for i, url in enumerate(urls):
        progress_val = (i + 1) / len(urls)
        progress_bar.progress(progress_val)
        status_text.write(f"Processing ({i+1}/{len(urls)}): `{url}`")
        
        logs.append(f"Checking URL: {url}")
        log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
        
        try:
            # Check if this is the CISA KEV Catalog URL
            if "known-exploited-vulnerabilities-catalog" in url.lower():
                logs.append(f"[*] CISA KEV Catalog URL detected. Fetching official KEV catalog JSON...")
                log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                kev_findings = parse_cisa_kev(url)
                results_list.extend(kev_findings)
                cves_found = [f['CVE ID'] for f in kev_findings if f['Published Today']]
                logs.append(f"<span style='color: #4ade80;'>[+] Parsed CISA KEV Catalog. Found {len(kev_findings)} recent entries ({len(cves_found)} added today/yesterday: {', '.join(cves_found)})</span>")
                log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                continue
                
            # Fetch content
            html = fetch_url(url)
            
            # Determine if this URL is a portal homepage/index page
            from urllib.parse import urlparse
            parsed = urlparse(url)
            path = parsed.path.strip('/')
            
            is_index = False
            if url in DEFAULT_SOURCES.values() or url.strip('/') in [v.strip('/') for v in DEFAULT_SOURCES.values()]:
                is_index = True
            elif not path or path in ['news', 'news/', 'alerts', 'alerts/', 'blog', 'blog/', 'updates', 'updates/', 'vuln', 'vuln/']:
                is_index = True
                
            if is_index:
                logs.append(f"[*] Portal/Index URL detected. Extracting article links...")
                log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                
                article_links = extract_article_links(url, html)
                # Limit to first 6 articles to ensure reasonable scan times and avoid abuse
                article_links = article_links[:6]
                
                if article_links:
                    logs.append(f"[+] Found {len(article_links)} latest articles to scan.")
                    log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                    
                    sub_findings_count = 0
                    for a_url in article_links:
                        status_text.write(f"Processing sub-article: `{a_url}`")
                        logs.append(f"Scanning sub-article: {a_url}")
                        log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                        try:
                            a_html = fetch_url(a_url)
                            a_findings = parse_article(a_url, a_html)
                            if a_findings:
                                results_list.extend(a_findings)
                                sub_findings_count += len(a_findings)
                                a_cves = [f['CVE ID'] for f in a_findings]
                                logs.append(f"<span style='color: #4ade80;'>&nbsp;&nbsp;[+] Found {len(a_findings)} CVEs: {', '.join(a_cves)}</span>")
                            else:
                                logs.append(f"<span style='color: #94a3b8;'>&nbsp;&nbsp;[-] No CVEs found.</span>")
                        except Exception as sub_e:
                            logs.append(f"<span style='color: #f87171;'>&nbsp;&nbsp;[!] Error: {str(sub_e)}</span>")
                        log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
                        time.sleep(0.2)
                    
                    logs.append(f"<span style='color: #3b82f6;'>[*] Index scan finished. Total CVEs: {sub_findings_count}</span>")
                else:
                    logs.append(f"<span style='color: #94a3b8;'>[-] No article links found on index page.</span>")
            else:
                # Treat as a direct article page
                findings = parse_article(url, html)
                if findings:
                    results_list.extend(findings)
                    cves_found = [f['CVE ID'] for f in findings]
                    logs.append(f"<span style='color: #4ade80;'>[+] Found {len(findings)} CVEs: {', '.join(cves_found)}</span>")
                else:
                    logs.append(f"<span style='color: #94a3b8;'>[-] No CVEs found in this article.</span>")
                
        except Exception as e:
            logs.append(f"<span style='color: #f87171;'>[!] Error: {str(e)}</span>")
            
        # Update logs container
        log_box.markdown(f'<div class="log-container">{"<br>".join(logs)}</div>', unsafe_allow_html=True)
        # Small sleep to simulate fluid real-time streaming & respect rate-limiting
        time.sleep(0.3)
        
    st.session_state['running'] = False
    st.session_state['scan_complete'] = True
    st.session_state['results'] = results_list
    progress_bar.progress(1.0)
    status_text.write("✨ Scan Complete!")
    
# Display findings if available
if 'scan_complete' in st.session_state and st.session_state['scan_complete']:
    findings = st.session_state['results']
    
    st.write("---")
    st.subheader("3. Investigation Findings Summary")
    
    # Report generation current date
    today_dt = datetime.now()
    today_str = today_dt.strftime("%Y-%m-%d")
    today_formatted = today_dt.strftime("%B %d, %Y")
    
    st.info(f"📅 **Daily Scan Date:** {today_formatted}")
    
    # Filter for today's CVEs checkbox (Enabled by default as requested)
    filter_today = st.checkbox("Only display and export CVEs published/updated today", value=True, 
                               help="Filter the report to only include CVEs found in articles published today.")
    
    filtered_findings = findings
    if filter_today:
        filtered_findings = [f for f in findings if f.get('Published Today', False)]
        
        # Display filtering feedback
        skipped = len(findings) - len(filtered_findings)
        if skipped > 0:
            st.success(f"🔍 Displaying **{len(filtered_findings)}** CVEs published today. (Filtered out {skipped} older vulnerability mentions).")
        elif len(filtered_findings) > 0:
            st.success(f"🔍 Displaying **{len(filtered_findings)}** CVEs published today.")
            
    if len(filtered_findings) == 0:
        if filter_today and len(findings) > 0:
            st.warning("No CVEs published today were found. Uncheck 'Only display and export CVEs published/updated today' to view older CVEs.")
        else:
            st.warning("No CVEs could be found in the scanned URLs.")
    else:
        # Display summary boxes
        s_col1, s_col2, s_col3 = st.columns(3)
        with s_col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(filtered_findings)}</div>
                <div class="metric-label">Total CVE Matches</div>
            </div>
            """, unsafe_allow_html=True)
        with s_col2:
            unique_cve_count = len(set(f['CVE ID'] for f in filtered_findings))
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{unique_cve_count}</div>
                <div class="metric-label">Unique CVEs</div>
            </div>
            """, unsafe_allow_html=True)
        with s_col3:
            unique_sources = len(set(f['Source'] for f in filtered_findings))
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{unique_sources}</div>
                <div class="metric-label">Affected Sources</div>
            </div>
            """, unsafe_allow_html=True)
            
        st.write("")
        
        # Group findings by Title and Source to merge multiple CVEs into a single cell
        grouped_findings = group_findings(filtered_findings)

        # Sort findings by date descending, newest at top. None values go to the bottom.
        def get_sort_key(item):
            dt = item.get('Date')
            if dt is None:
                return datetime.min
            if isinstance(dt, str):
                try:
                    return datetime.strptime(dt, "%Y-%m-%d")
                except ValueError:
                    return datetime.min
            return dt

        grouped_findings = sorted(grouped_findings, key=get_sort_key, reverse=True)

        # Prepare excel data (excluding Date tags so they look neat in Excel)
        excel_data_list = []
        for f in grouped_findings:
            excel_data_list.append({
                'CVE ID': f['CVE ID'],
                'Severity': f['Severity'],
                'Teams Affected': f['Teams Affected'],
                'Title': f['Title'],
                'Source': f['Source'],
                'Description': f['Description']
            })
            
        # Download button with dynamic date filename
        excel_data = generate_excel_report(excel_data_list)
        report_filename = f"threat_intel_cve_report_{today_str}.xlsx"
        
        st.download_button(
            label=f"📥 Download Excel Report ({report_filename})",
            data=excel_data,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Dataframe preview
        st.write("**Report Preview:**")
        df_preview = pd.DataFrame(excel_data_list)
        df_preview.insert(0, 'S.No', range(1, len(df_preview) + 1))
        st.dataframe(
            df_preview,
            column_config={
                "S.No": st.column_config.NumberColumn("S.No", width="small"),
                "CVE ID": st.column_config.TextColumn("CVE ID", width="medium"),
                "Severity": st.column_config.TextColumn("Severity", width="small"),
                "Teams Affected": st.column_config.TextColumn("Teams Affected", width="medium"),
                "Title": st.column_config.TextColumn("Title of Article", width="large"),
                "Source": st.column_config.LinkColumn("Source URL", width="large"),
                "Description": st.column_config.TextColumn("Description", width="max"),
            },
            hide_index=True,
            width="stretch"
        )
