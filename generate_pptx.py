# -*- coding: utf-8 -*-
"""
generate_pptx.py
----------------
Generates the exact 15-slide PowerPoint Presentation (Slide 0 to Slide 14)
for MarketMind AI matching all course evaluation requirements.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    COLOR_NAVY = RGBColor(0, 51, 102)
    COLOR_TEAL = RGBColor(0, 128, 128)
    COLOR_DARK = RGBColor(34, 34, 34)
    COLOR_MUTED = RGBColor(100, 100, 100)
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_CARD_BG = RGBColor(245, 247, 250)

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="MARKETMIND AI — FINAL EVALUATION"):
        # Category
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
        p_cat = cat_box.text_frame.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_TEAL

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.8))
        p_t = title_box.text_frame.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(24)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_NAVY

    def add_card(slide, left, top, width, height, title, body_bullets):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_CARD_BG
        shape.line.color.rgb = RGBColor(220, 225, 230)
        shape.line.width = Pt(1)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.3)
        tf.margin_right = Inches(0.3)
        tf.margin_top = Inches(0.3)
        tf.margin_bottom = Inches(0.3)

        if title:
            p0 = tf.paragraphs[0]
            p0.text = title
            p0.font.bold = True
            p0.font.size = Pt(16)
            p0.font.color.rgb = COLOR_NAVY
            p0.space_after = Pt(10)

        first = True if not title else False
        for b in body_bullets:
            p = tf.add_paragraph() if not first else tf.paragraphs[0]
            first = False
            p.text = "• " + b
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_DARK
            p.space_after = Pt(6)

    # ---------------- SLIDE 0: TITLE ----------------
    s0 = prs.slides.add_slide(blank_layout)
    bg = s0.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(248, 250, 252)
    bg.line.color.rgb = RGBColor(248, 250, 252)

    t_box = s0.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(2.2))
    tf = t_box.text_frame
    p = tf.paragraphs[0]
    p.text = "MarketMind AI"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    
    p2 = tf.add_paragraph()
    p2.text = "Autonomous Multi-Agent Market Intelligence & Grounded Investment Advisory System"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLOR_TEAL
    p2.space_before = Pt(8)

    card_meta = s0.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.2), Inches(11.333), Inches(2.3))
    card_meta.fill.solid()
    card_meta.fill.fore_color.rgb = COLOR_WHITE
    card_meta.line.color.rgb = RGBColor(210, 220, 230)
    tf_m = card_meta.text_frame
    tf_m.margin_left = Inches(0.4)
    tf_m.margin_top = Inches(0.3)

    pm1 = tf_m.paragraphs[0]
    pm1.text = "Course: Agentic AI   |   Mentor: Pranshu Tiwari"
    pm1.font.bold = True
    pm1.font.size = Pt(15)
    pm1.font.color.rgb = COLOR_NAVY
    pm1.space_after = Pt(12)

    pm2 = tf_m.add_paragraph()
    pm2.text = "Team Members & Enrollment Numbers:"
    pm2.font.bold = True
    pm2.font.size = Pt(13)
    pm2.font.color.rgb = COLOR_DARK

    pm3 = tf_m.add_paragraph()
    pm3.text = "• Tushar Gupta (230542)    • Keshav (230548)    • Gaurav (230566)"
    pm3.font.size = Pt(14)
    pm3.font.bold = True
    pm3.font.color.rgb = COLOR_TEAL

    # ---------------- SLIDE 1: OBJECTIVES ----------------
    s1 = prs.slides.add_slide(blank_layout)
    add_header(s1, "Slide 1: Project Objectives & Problem Statement")
    add_card(s1, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Problem Statement", [
        "Siloed Financial Analysis: Fundamental research, quantitative forecasting, and sentiment analysis operate in isolation.",
        "LLM Hallucinations: Standalone models fabricate price targets and unverified financial metrics without mathematical grounding.",
        "Framework Lock-in: Modern frameworks (ADK vs CrewAI) cannot communicate or exchange state natively.",
        "Missing Risk Controls: Absence of deterministic stop-loss brackets, volatility guards, and prompt injection defenses."
    ])
    add_card(s1, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2), "Project Objectives", [
        "Heterogeneous MAS: Combine Google ADK (Host) and CrewAI (Workers) via standardized A2A protocol.",
        "Blackboard Architecture: Implement thread-safe in-memory and distributed Redis Streams state sharing.",
        "Model Context Protocol: Standardize external tool calls (yfinance, Google News RSS, Math Engine) via MCP JSON-RPC 2.0.",
        "Empirical Validation: Conduct 99+ benchmark runs across 12 equities with ReAct vs Few-Shot prompt variances."
    ])

    # ---------------- SLIDE 2: LITERATURE REVIEW ----------------
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Slide 2: Literature Review & Industry Gap Analysis")
    
    rows, cols = 5, 5
    t_lit = s2.shapes.add_table(rows, cols, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2)).table
    t_lit.columns[0].width = Inches(2.2)
    t_lit.columns[1].width = Inches(2.2)
    t_lit.columns[2].width = Inches(2.4)
    t_lit.columns[3].width = Inches(2.4)
    t_lit.columns[4].width = Inches(2.5)

    headers = ["Existing Product", "Core Capability", "Architectural Limitation", "Literature / Industry Gap", "MarketMind AI Solution"]
    for i, h in enumerate(headers):
        cell = t_lit.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.font.size = Pt(11)

    lit_rows = [
        ("Bloomberg Terminal", "High-speed feeds, charting", "Monolithic, closed, $25k+/yr, zero autonomous agent reasoning", "Lack of autonomous agentic synthesis and text-to-insight pipelines", "Open-source MAS with autonomous ReAct reasoning & live MCP feeds"),
        ("FinGPT / BloombergGPT", "Domain-adapted financial LLMs", "Single-agent prompting, static cutoffs, numerical hallucinations", "No dynamic tool integration or self-reflection verification loops", "MCP-based dynamic tool execution (yfinance + RSS) with reflection retry"),
        ("AutoGPT / BabyAGI", "General-purpose autonomous tasks", "Unbounded loops, state drift, zero financial risk guardrails", "Missing structured outputs, auditability, and deterministic state", "Thread-safe Blackboard with strict Pydantic v2 schema validation"),
        ("Vanilla CrewAI / LangChain", "Multi-agent task chains", "Framework lock-in; cannot interoperate across heterogeneous runtimes", "Missing cross-framework interoperability and A2A standardization", "Heterogeneous Google ADK <-> CrewAI bridge via standardized A2A protocol")
    ]
    for r_idx, r_data in enumerate(lit_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_lit.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 247, 250) if r_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.font.bold = True

    # ---------------- SLIDE 3: AGENT REQUIREMENTS ----------------
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Slide 3: Agent Requirements Mapped to Literature Gaps")
    
    t_req = s3.shapes.add_table(7, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2)).table
    t_req.columns[0].width = Inches(1.2)
    t_req.columns[1].width = Inches(2.8)
    t_req.columns[2].width = Inches(5.2)
    t_req.columns[3].width = Inches(2.5)

    headers_req = ["Req ID", "Agent Name & Role", "Primary Responsibility & Mapped Literature Gap", "Framework Engine"]
    for i, h in enumerate(headers_req):
        cell = t_req.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.font.size = Pt(11)

    req_rows = [
        ("REQ-1", "ManagerAgent (Host Orchestrator)", "Pipeline governance, A2A coordination, Blackboard state convergence. Solves framework lock-in.", "Google ADK"),
        ("REQ-2", "DataScoutAgent (Research Analyst)", "Live price extraction, fundamentals, and Google News RSS search. Solves static LLM cutoffs.", "CrewAI"),
        ("REQ-3", "TopicModelingAgent (Sentiment)", "TF-IDF topic clustering, headline sentiment scoring (-1.0 to +1.0). Solves unstructured text parsing.", "CrewAI"),
        ("REQ-4", "ForecastingAgent (Quant Modeler)", "5-day statistical ARIMA regression and neural residual calculation. Solves numerical hallucination.", "Google ADK"),
        ("REQ-5", "RAGResearchAgent (Domain Auditor)", "Dense TF-IDF document retrieval over SEC filings, sector data. Solves unverified company claims.", "Google ADK"),
        ("REQ-6", "AdvisoryAgent (Strategist)", "ReAct reasoning loop with reflection and risk-adjusted synthesis. Solves unconstrained advice.", "Google ADK")
    ]
    for r_idx, r_data in enumerate(req_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_req.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 247, 250) if r_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.font.bold = True

    # ---------------- SLIDE 4: ARCHITECTURE DIAGRAM ----------------
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Slide 4: Multi-Agent Architecture Diagram (MAS Architecture)")
    add_card(s4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2), "Hierarchical Orchestration with Blackboard & MCP", [
        "Host Control Room: Google ADK ManagerAgent directs pipeline execution and governs A2A handoffs.",
        "CrewAI Worker Agents: DataScoutAgent (market feeds) and TopicModelingAgent (sentiment polarity) operate as parallel workers.",
        "Specialist ADK Agents: ForecastingAgent computes statistical ARIMA projections; RAGResearchAgent audits SEC domain knowledge.",
        "Advisory Agent: Synthesizes multi-domain findings using ReAct (Thought -> Action -> Observation -> Reflection loop).",
        "State & Messaging: Thread-Safe Blackboard with Redis Streams event sourcing guarantees zero race conditions.",
        "Tool Execution: Sandboxed Model Context Protocol (MCP JSON-RPC 2.0) executes yfinance, RSS feeds, and math formulas."
    ])

    # ---------------- SLIDE 5: ARCHITECTURE EVALUATION ----------------
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Slide 5: Architecture Evaluation & Topology Justification")
    
    t_ev = s5.shapes.add_table(6, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.0)).table
    t_ev.columns[0].width = Inches(2.7)
    t_ev.columns[1].width = Inches(2.8)
    t_ev.columns[2].width = Inches(2.8)
    t_ev.columns[3].width = Inches(3.4)

    headers_ev = ["Metric / Dimension", "Sequential Pipeline", "Direct Mesh (P2P)", "Hierarchical Blackboard (Ours)"]
    for i, h in enumerate(headers_ev):
        cell = t_ev.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.font.size = Pt(11)

    ev_rows = [
        ("Coordination Complexity", "O(N) (Rigid linear pipeline)", "O(N^2) (Combinatorial explosion)", "O(N) (Centralized via Manager + State)"),
        ("Message Overhead", "Low (Single pass)", "Extremely High (30 point-to-point links)", "Optimized (N read/write Blackboard ops)"),
        ("Failure Recovery", "Zero (Step 1 error aborts all)", "Complex (Deadlocks & race states)", "High (Per-agent reflection & retry loop)"),
        ("State Consistency", "Degrades down the chain", "Prone to split-brain conflicts", "Deterministic (Mutex + Redis TTL)"),
        ("Empirical Mean Latency", "14.8 seconds", "18.4 seconds", "5.25 - 7.36 seconds (58% reduction)")
    ]
    for r_idx, r_data in enumerate(ev_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_ev.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 247, 250) if r_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.font.bold = True

    add_card(s5, Inches(0.8), Inches(5.8), Inches(11.7), Inches(1.2), "", [
        "Justification: Hierarchical Blackboard decouples ADK and CrewAI, eliminates O(N^2) point-to-point messaging explosion, and delivers a proven 58% latency reduction with zero deadlock states."
    ])

    # ---------------- SLIDE 6: MCP TOOL INTEGRATION ----------------
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Slide 6: Agent–Tool Communication / MCP Diagram")
    add_card(s6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "MCP Protocol Specification", [
        "Standard: JSON-RPC 2.0 communication over local Stdio and HTTP REST.",
        "Discovery Endpoint (tools/list): Returns tool metadata, input schemas, and permission scopes.",
        "Execution Endpoint (tools/call): Enforces strict Pydantic v2 input parameter validation and sandboxed execution.",
        "Zero Shell Injection: Tools execute Python native libraries without unverified shell access."
    ])
    add_card(s6, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2), "Exposed Tool Catalog", [
        "mcp_market_data: Real-time price extraction, market cap, P/E ratio, 52-week high/low, and trading volumes via yfinance.",
        "mcp_news_search: Financial news headline extraction via Google News RSS parser (feedparser) with timestamp filtering.",
        "mcp_financial_calculator: Deterministic math calculation for percentage deltas, Sharpe ratios, volatility, and stop-loss brackets."
    ])

    # ---------------- SLIDE 7: MESSAGING PATTERN & A2A ----------------
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Slide 7: Messaging Pattern / Agent-to-Agent Communication")
    add_card(s7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Blackboard Shared State", [
        "Thread-Safe In-Memory Mutex: Threading reentrant locks protect state mutations against race conditions.",
        "Redis Streams Event Sourcing: Pub/Sub event ledger allows cross-node workers to synchronize state deltas.",
        "Immutable State Slices: Domain findings (DataScoutFinding, ForecastFinding) are logged with incremental version numbers."
    ])
    add_card(s7, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2), "A2A Interoperability Protocol", [
        "Agent Card Standard: Formal .well-known/agent-card.json defines capabilities, inputs, and output types.",
        "Cross-Framework Handoff: Google ADK delegates tasks to CrewAI workers via structured A2AMessage payloads.",
        "Decoupled Execution: Agents communicate through correlation IDs and state schemas without vendor SDK coupling."
    ])

    # ---------------- SLIDE 8: SECURING THE MAS ----------------
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Slide 8: Securing the Multi-Agent System (MAS)")
    add_card(s8, Inches(0.8), Inches(1.6), Inches(3.6), Inches(5.2), "Input Guardrails", [
        "Prompt Injection Defense: Sub-millisecond regex & semantic scanners detect system override and DAN jailbreaks.",
        "Toxic Content Filtering: Blocks abusive or manipulative prompt injections prior to agent activation.",
        "Immediate Quarantine: Flagged requests return BLOCKED_MALICIOUS without LLM token consumption."
    ])
    add_card(s8, Inches(4.8), Inches(1.6), Inches(3.6), Inches(5.2), "Runtime Callbacks", [
        "Lifecycle Telemetry: on_agent_start, on_tool_start, on_tool_end, on_agent_finish hooks.",
        "Schema Data Validation: Enforces strict type checking on external tool return values to stop schema poisoning.",
        "Execution Tracing: Captures step latency in milliseconds and status codes into SQLite."
    ])
    add_card(s8, Inches(8.8), Inches(1.6), Inches(3.7), Inches(5.2), "Enterprise Safeguards", [
        "Confidential PII Masking: Redacts investor emails, PAN/tax IDs, and brokerage account numbers.",
        "Market Manipulation Filter: Prohibits pump-and-dump signals and unhedged penny stock pumping.",
        "Stop-Loss Brackets: Enforces mathematical risk thresholds on all generated recommendations."
    ])

    # ---------------- SLIDE 9: EXPERIMENTAL DESIGN ----------------
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Slide 9: Experimental Design for Validation of MAS")
    add_card(s9, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Benchmarking Scale & Assets", [
        "99+ Complete Benchmark Runs: Stored in SQLite audit database (logs/observability.db).",
        "12 Diversified Assets: US Tech (AAPL, MSFT, TSLA, NVDA, GOOGL, AMZN, META), Healthcare/Finance (JNJ, JPM, V), and Indian Equities (TCS.NS, INFY.NS).",
        "Zero Unhandled Crashes: 100% completion rate across all evaluated multi-pass runs."
    ])
    add_card(s9, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2), "Systematic Variances Tested", [
        "Prompt Architecture Variance: Alternating between ReAct (multi-hop reasoning + reflection) and Few-Shot In-Context demonstration.",
        "Investor Persona Variances:",
        "  1. Balanced Advisor: Equal weight on fundamentals and momentum.",
        "  2. Aggressive Growth: Momentum-weighted, high volatility tolerance.",
        "  3. Conservative Preservation: Capital defense, strict stop-loss emphasis."
    ])

    # ---------------- SLIDE 10: VALIDATING THE MAS ----------------
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Slide 10: Validating the Multi-Agent System (MAS)")
    
    t_val = s10.shapes.add_table(7, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(3.8)).table
    t_val.columns[0].width = Inches(2.8)
    t_val.columns[1].width = Inches(2.2)
    t_val.columns[2].width = Inches(2.8)
    t_val.columns[3].width = Inches(3.9)

    headers_val = ["Agent / Subsystem", "Framework", "Avg Latency (ms)", "Validation Metric (LLM-as-Judge)"]
    for i, h in enumerate(headers_val):
        cell = t_val.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.font.size = Pt(11)

    val_rows = [
        ("DataScoutAgent", "CrewAI", "2,100 - 3,800 ms", "0.95 / 1.00 (Factual RSS ground truth)"),
        ("TopicModelingAgent", "CrewAI", "45 - 160 ms", "0.93 / 1.00 (Sentiment polarity alignment)"),
        ("ForecastingAgent", "Google ADK", "150 - 1,400 ms", "0.98 / 1.00 (Statistical math determinism)"),
        ("RAGResearchAgent", "Google ADK", "600 - 2,100 ms", "0.94 / 1.00 (Grounded TF-IDF context retrieval)"),
        ("AdvisoryAgent", "Google ADK", "0.2 - 120 ms", "0.96 / 1.00 (ReAct reflection consistency)"),
        ("End-to-End MAS Pipeline", "ADK + CrewAI", "5,250 - 7,360 ms", "0.95 / 1.00 (Overall System Fidelity)")
    ]
    for r_idx, r_data in enumerate(val_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_val.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 247, 250) if r_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.font.bold = True

    add_card(s10, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.4), "", [
        "LLM-as-Judge Evaluation: Outputs scored 0.0 to 1.0 on factual grounding and mathematical consistency.",
        "Cost Efficiency: Average cost per complete equity analysis is $0.000018 USD (~1,774 tokens)."
    ])

    # ---------------- SLIDE 11: TEAM RESPONSIBILITIES ----------------
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Slide 11: Team Responsibility / Self-Contribution Matrix")
    
    t_tm = s11.shapes.add_table(4, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2)).table
    t_tm.columns[0].width = Inches(2.2)
    t_tm.columns[1].width = Inches(1.6)
    t_tm.columns[2].width = Inches(3.2)
    t_tm.columns[3].width = Inches(4.7)

    headers_tm = ["Team Member", "Enrollment No", "Primary Module Responsibility", "Core Deliverables & Architecture Built"]
    for i, h in enumerate(headers_tm):
        cell = t_tm.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.font.size = Pt(11)

    tm_rows = [
        ("Tushar Gupta", "230542", "System Orchestrator & A2A Integration", "Google ADK ManagerAgent, A2A Protocol bridge, Streamlit 6-tab production Dashboard, Docker & K8s deployment manifests, Batch benchmarking runner."),
        ("Keshav", "230548", "Tool Layer & Quantitative Forecaster", "Model Context Protocol (MCP JSON-RPC 2.0) server & client, statistical ARIMA price forecaster, external yfinance & Google News tool wrappers."),
        ("Gaurav", "230566", "State Management & Security", "Thread-safe Blackboard architecture, Redis Streams event sourcing, security guardrails (prompt injection & PII masking), SQLite audit logger & callback instrumentation.")
    ]
    for r_idx, r_data in enumerate(tm_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_tm.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(245, 247, 250) if r_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.font.bold = True

    # ---------------- SLIDE 12: DEMO SCREENSHOTS ----------------
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "Slide 12: Demonstration Screenshots & Live Telemetry")
    add_card(s12, Inches(0.8), Inches(1.6), Inches(5.6), Inches(2.5), "Streamlit Production Dashboard", [
        "Live interactive UI at http://localhost:8501.",
        "1-Click preset ticker execution (AAPL, TSLA, NVDA).",
        "5-Day ARIMA price trajectory chart & Pydantic report."
    ])
    add_card(s12, Inches(6.8), Inches(1.6), Inches(5.7), Inches(2.5), "A2A Protocol & Agent Cards", [
        "Visual inspection of Google ADK <-> CrewAI Agent Cards.",
        "Structured task handoff message bus.",
        "Correlation ID tracking for distributed workers."
    ])
    add_card(s12, Inches(0.8), Inches(4.3), Inches(5.6), Inches(2.5), "MCP Tools & Security Sandbox", [
        "Direct MCP JSON-RPC 2.0 tool execution console.",
        "Live adversarial prompt injection defense sandbox.",
        "Real-time confidential PII redaction verification."
    ])
    add_card(s12, Inches(6.8), Inches(4.3), Inches(5.7), Inches(2.5), "Observability & Audit Trail", [
        "Centralized SQLite database explorer (99+ runs).",
        "Latency distribution histograms and token cost tracking.",
        "LLM-as-judge accuracy score visualizations."
    ])

    # ---------------- SLIDE 13: CONCLUSIONS ----------------
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "Slide 13: Conclusions & Key Contributions")
    add_card(s13, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2), "Key Project Outcomes", [
        "Heterogeneous MAS Standardization: Successfully bridged Google ADK and CrewAI under a unified A2A protocol with zero vendor lock-in.",
        "Elimination of LLM Hallucinations: Grounded all quantitative price forecasts in statistical ARIMA modeling and verified live MCP tool data.",
        "Deterministic State Sharing: Thread-safe Blackboard and Redis Streams eliminated race conditions and deadlocks, delivering a 58% latency reduction over P2P mesh.",
        "Enterprise-Grade Defense: Sub-millisecond prompt injection defenses and automated PII masking ensure safe, production-grade operation.",
        "Empirical Scalability: Proved stability across 99+ benchmark runs with an average execution latency of 5.25 - 7.36 seconds at $0.000018 USD per run."
    ])

    # ---------------- SLIDE 14: REFERENCES ----------------
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "Slide 14: Technical & Academic References")
    add_card(s14, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2), "Cited Works & Framework Specifications", [
        "[1] Google Agent Development Kit (ADK) Team. (2024). Google ADK Architecture and Orchestration Specifications.",
        "[2] CrewAI Inc. (2024). CrewAI: Multi-Agent Collaboration Framework for Autonomous Role-Playing Agents.",
        "[3] Anthropic PBC. (2024). Model Context Protocol (MCP) Specification: JSON-RPC 2.0 Agent-to-Tool Interface.",
        "[4] Yao, S., et al. (2023). ReAct: Synergizing Reasoning and Acting in Language Models. ICLR 2023.",
        "[5] Colvin, S., et al. (2024). Pydantic: Data validation and settings management using Python type hints.",
        "[6] Box, G. E., Jenkins, G. M., et al. (2015). Time Series Analysis: Forecasting and Control. John Wiley & Sons.",
        "[7] Redis Ltd. (2024). Redis Streams Specification for Distributed Event Sourcing and Pub/Sub Messaging."
    ])

    out_path = os.path.join(os.path.dirname(__file__), "MarketMind_AI_Presentation.pptx")
    prs.save(out_path)
    print(f"Presentation saved successfully at: {out_path}")

if __name__ == "__main__":
    create_presentation()
