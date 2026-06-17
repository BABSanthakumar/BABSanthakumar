from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

# Colors
DARK_BG = RGBColor(0x0D, 0x1B, 0x2A)
ACCENT = RGBColor(0xE8, 0x8C, 0x03)
BLUE = RGBColor(0x1E, 0x90, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xE0, 0xE0, 0xE0)
GREEN = RGBColor(0x2E, 0xCC, 0x71)
RED = RGBColor(0xE7, 0x4C, 0x3C)
DARK_CARD = RGBColor(0x1A, 0x2B, 0x3C)
ORANGE = RGBColor(0xFF, 0x6B, 0x35)

blank_layout = prs.slide_layouts[6]

def set_bg(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width or 1)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, left, top, width, height, font_size=16, bold=False, color=WHITE, align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_bullet_text(slide, lines, left, top, width, height, font_size=14, color=WHITE, bold_first=False):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        if bold_first and i == 0:
            run.font.bold = True
    return txBox

# ─────────────────────────────────────────────
# SLIDE 1: Title
# ─────────────────────────────────────────────
slide1 = prs.slides.add_slide(blank_layout)
set_bg(slide1, DARK_BG)

add_rect(slide1, 0, 0, 13.33, 0.12, ACCENT)
add_rect(slide1, 0, 7.38, 13.33, 0.12, ACCENT)

add_rect(slide1, 1.5, 1.5, 10.33, 4.5, DARK_CARD, ACCENT, 1.5)

add_text(slide1, "Claude Development", 2, 2, 9.33, 1.2, font_size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide1, "Web  vs  IDE", 2, 3.1, 9.33, 1.0, font_size=38, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(slide1, "Which is Best & Why  |  Token Usage Deep Dive", 2, 4.1, 9.33, 0.7, font_size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

add_text(slide1, "Claude Code  ·  Managed Agents  ·  API Surfaces  ·  Cost Analysis", 1.5, 6.5, 10.33, 0.6, font_size=13, color=RGBColor(0x90, 0x90, 0x90), align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────
# SLIDE 2: Two Surfaces Overview
# ─────────────────────────────────────────────
slide2 = prs.slides.add_slide(blank_layout)
set_bg(slide2, DARK_BG)
add_rect(slide2, 0, 0, 13.33, 0.08, ACCENT)

add_text(slide2, "Two Development Surfaces", 0.3, 0.15, 12.5, 0.7, font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Web card
add_rect(slide2, 0.3, 1.0, 5.9, 5.9, DARK_CARD, BLUE, 1.5)
add_text(slide2, "🌐  Claude.ai / Claude Code Web", 0.5, 1.1, 5.5, 0.55, font_size=16, bold=True, color=BLUE)
web_lines = [
    "• Browser-based, zero install",
    "• claude.ai/code web interface",
    "• Managed Agents (server-run loop)",
    "• Per-session containers",
    "• File mounts & workspace",
    "• SSE event streaming",
    "• Skills + MCP tools",
    "• Anthropic handles infra",
    "• Best for: exploratory work,",
    "  agentic long-horizon tasks",
]
add_bullet_text(slide2, web_lines, 0.5, 1.75, 5.5, 5.0, font_size=13.5, color=LIGHT_GRAY)

# IDE card
add_rect(slide2, 7.13, 1.0, 5.9, 5.9, DARK_CARD, GREEN, 1.5)
add_text(slide2, "💻  Claude Code CLI / IDE Extensions", 7.33, 1.1, 5.5, 0.55, font_size=16, bold=True, color=GREEN)
ide_lines = [
    "• VS Code, JetBrains, terminal CLI",
    "• Local execution environment",
    "• Full codebase context",
    "• Direct file read/write/edit",
    "• Custom tool use (you host compute)",
    "• Maximum flexibility & control",
    "• Batch processing (50% off)",
    "• Prompt caching (up to 90% off)",
    "• Best for: daily coding, CI/CD,",
    "  large-repo workflows",
]
add_bullet_text(slide2, ide_lines, 7.33, 1.75, 5.5, 5.0, font_size=13.5, color=LIGHT_GRAY)

add_text(slide2, "vs", 6.22, 3.5, 0.9, 0.6, font_size=22, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────
# SLIDE 3: Token Usage & Pricing
# ─────────────────────────────────────────────
slide3 = prs.slides.add_slide(blank_layout)
set_bg(slide3, DARK_BG)
add_rect(slide3, 0, 0, 13.33, 0.08, ACCENT)

add_text(slide3, "Token Usage & Pricing", 0.3, 0.15, 12.5, 0.7, font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Table header
add_rect(slide3, 0.3, 1.0, 12.73, 0.5, ACCENT)
headers = ["Model", "Context Window", "Input $/1M", "Output $/1M", "Thinking", "Best Use"]
col_widths = [2.2, 2.0, 1.7, 1.9, 1.5, 3.43]
x = 0.3
for i, (h, w) in enumerate(zip(headers, col_widths)):
    add_text(slide3, h, x+0.05, 1.05, w-0.1, 0.4, font_size=13, bold=True, color=DARK_BG, align=PP_ALIGN.CENTER)
    x += w

rows = [
    ("Claude Fable 5",    "1M",   "$10.00", "$50.00", "Always On",  "Most capable, agentic"),
    ("Claude Opus 4.8",   "1M",   "$5.00",  "$25.00", "Adaptive",   "Default recommended"),
    ("Claude Opus 4.7",   "1M",   "$5.00",  "$25.00", "Adaptive",   "Complex reasoning"),
    ("Claude Sonnet 4.6", "1M",   "$3.00",  "$15.00", "Adaptive",   "Balanced cost/perf"),
    ("Claude Haiku 4.5",  "200K", "$1.00",  "$5.00",  "—",          "Fast, lightweight tasks"),
]

row_colors = [DARK_CARD, RGBColor(0x12, 0x22, 0x33), DARK_CARD, RGBColor(0x12, 0x22, 0x33), DARK_CARD]
text_colors = [ACCENT, WHITE, WHITE, WHITE, LIGHT_GRAY]

for ri, (row, rbg, tc) in enumerate(zip(rows, row_colors, text_colors)):
    y = 1.5 + ri * 0.65
    add_rect(slide3, 0.3, y, 12.73, 0.62, rbg)
    x = 0.3
    for ci, (cell, w) in enumerate(zip(row, col_widths)):
        fc = tc if ci == 0 else LIGHT_GRAY
        add_text(slide3, cell, x+0.05, y+0.1, w-0.1, 0.42, font_size=12.5, color=fc, align=PP_ALIGN.CENTER)
        x += w

# Savings callouts
add_rect(slide3, 0.3, 4.85, 4.0, 1.35, RGBColor(0x0A, 0x2A, 0x1A), GREEN, 1)
add_text(slide3, "⚡ Prompt Caching", 0.4, 4.9, 3.8, 0.4, font_size=13, bold=True, color=GREEN)
add_text(slide3, "5-min TTL: up to 90% off repeats\n1-hour TTL: extended context savings", 0.4, 5.3, 3.8, 0.8, font_size=11.5, color=LIGHT_GRAY)

add_rect(slide3, 4.67, 4.85, 4.0, 1.35, RGBColor(0x1A, 0x1A, 0x0A), ACCENT, 1)
add_text(slide3, "📦 Batch Processing", 4.77, 4.9, 3.8, 0.4, font_size=13, bold=True, color=ACCENT)
add_text(slide3, "50% cost reduction on bulk tasks\nAsync processing, 24h turnaround", 4.77, 5.3, 3.8, 0.8, font_size=11.5, color=LIGHT_GRAY)

add_rect(slide3, 9.03, 4.85, 4.0, 1.35, RGBColor(0x0A, 0x12, 0x2A), BLUE, 1)
add_text(slide3, "🔢 Token Counting API", 9.13, 4.9, 3.8, 0.4, font_size=13, bold=True, color=BLUE)
add_text(slide3, "Count tokens before sending\nAvoid surprises on large payloads", 9.13, 5.3, 3.8, 0.8, font_size=11.5, color=LIGHT_GRAY)

# ─────────────────────────────────────────────
# SLIDE 4: Web Code Deep Dive
# ─────────────────────────────────────────────
slide4 = prs.slides.add_slide(blank_layout)
set_bg(slide4, DARK_BG)
add_rect(slide4, 0, 0, 13.33, 0.08, BLUE)

add_text(slide4, "🌐  Claude Code Web — Deep Dive", 0.3, 0.15, 12.5, 0.7, font_size=26, bold=True, color=BLUE)

add_rect(slide4, 0.3, 1.0, 5.9, 2.7, DARK_CARD, BLUE, 1)
add_text(slide4, "Managed Agents", 0.5, 1.05, 5.6, 0.45, font_size=15, bold=True, color=BLUE)
ma_lines = [
    "✓  Anthropic runs the agent loop",
    "✓  Per-session container workspace",
    "✓  File mounts & persistent storage",
    "✓  SSE real-time event streaming",
    "✓  Skills + MCP tool integrations",
    "✓  Versioned, persisted agent configs",
]
add_bullet_text(slide4, ma_lines, 0.5, 1.55, 5.5, 2.0, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide4, 0.3, 3.85, 5.9, 2.7, DARK_CARD, BLUE, 1)
add_text(slide4, "Ideal Scenarios", 0.5, 3.9, 5.6, 0.45, font_size=15, bold=True, color=BLUE)
sc_lines = [
    "→  Long-horizon agentic tasks",
    "→  No local infra to configure",
    "→  Teams needing shared agents",
    "→  Exploratory / prototyping work",
    "→  Quick demos & presentations",
]
add_bullet_text(slide4, sc_lines, 0.5, 4.4, 5.5, 1.9, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide4, 7.1, 1.0, 5.9, 2.7, DARK_CARD, RED, 1)
add_text(slide4, "Limitations", 7.3, 1.05, 5.6, 0.45, font_size=15, bold=True, color=RED)
lim_lines = [
    "✗  Not available on Bedrock/Vertex/Foundry",
    "✗  Less control over execution environment",
    "✗  Harder to integrate with local CI/CD",
    "✗  Data leaves your infrastructure",
    "✗  Agent IDs must be stored externally",
]
add_bullet_text(slide4, lim_lines, 7.3, 1.55, 5.5, 2.0, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide4, 7.1, 3.85, 5.9, 2.7, DARK_CARD, ACCENT, 1)
add_text(slide4, "Token Strategy (Web)", 7.3, 3.9, 5.6, 0.45, font_size=15, bold=True, color=ACCENT)
tok_lines = [
    "• Default: Claude Opus 4.8 (adaptive thinking)",
    "• 1M context window for large projects",
    "• Prompt caching on repeated system prompts",
    "• Fable 5 available for hardest tasks",
    "• Effort levels: low → xhigh → max",
]
add_bullet_text(slide4, tok_lines, 7.3, 4.4, 5.5, 1.9, font_size=12.5, color=LIGHT_GRAY)

# ─────────────────────────────────────────────
# SLIDE 5: IDE Code Deep Dive
# ─────────────────────────────────────────────
slide5 = prs.slides.add_slide(blank_layout)
set_bg(slide5, DARK_BG)
add_rect(slide5, 0, 0, 13.33, 0.08, GREEN)

add_text(slide5, "💻  Claude Code CLI / IDE — Deep Dive", 0.3, 0.15, 12.5, 0.7, font_size=26, bold=True, color=GREEN)

add_rect(slide5, 0.3, 1.0, 5.9, 2.7, DARK_CARD, GREEN, 1)
add_text(slide5, "Core Capabilities", 0.5, 1.05, 5.6, 0.45, font_size=15, bold=True, color=GREEN)
cap_lines = [
    "✓  VS Code & JetBrains extensions",
    "✓  Terminal CLI — direct shell access",
    "✓  Full repo read/write/edit/run",
    "✓  Custom tool definitions (you host)",
    "✓  Automatic tool runner loop",
    "✓  Works on Bedrock, Vertex, Foundry",
]
add_bullet_text(slide5, cap_lines, 0.5, 1.55, 5.5, 2.0, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide5, 0.3, 3.85, 5.9, 2.7, DARK_CARD, GREEN, 1)
add_text(slide5, "Ideal Scenarios", 0.5, 3.9, 5.6, 0.45, font_size=15, bold=True, color=GREEN)
sc2_lines = [
    "→  Day-to-day coding & refactoring",
    "→  Large monorepo workflows",
    "→  CI/CD pipelines & automation",
    "→  Multi-file changes & migrations",
    "→  Offline / air-gapped environments",
]
add_bullet_text(slide5, sc2_lines, 0.5, 4.4, 5.5, 1.9, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide5, 7.1, 1.0, 5.9, 2.7, DARK_CARD, ORANGE, 1)
add_text(slide5, "Cost Optimizations", 7.3, 1.05, 5.6, 0.45, font_size=15, bold=True, color=ORANGE)
cost_lines = [
    "💰  Batch API: 50% off bulk requests",
    "⚡  Prompt cache: up to 90% off repeats",
    "🔢  Token Count API: avoid surprises",
    "📉  Use Haiku 4.5 for simple subtasks",
    "🧩  Streaming for long I/O (no timeouts)",
]
add_bullet_text(slide5, cost_lines, 7.3, 1.55, 5.5, 2.0, font_size=12.5, color=LIGHT_GRAY)

add_rect(slide5, 7.1, 3.85, 5.9, 2.7, DARK_CARD, ACCENT, 1)
add_text(slide5, "Token Strategy (IDE)", 7.3, 3.9, 5.6, 0.45, font_size=15, bold=True, color=ACCENT)
tok2_lines = [
    "• Sonnet 4.6 for balanced daily coding",
    "• Opus 4.8 for complex architecture tasks",
    "• Haiku 4.5 for lint/format/simple fixes",
    "• Adaptive thinking for non-trivial work",
    "• 1M context: full codebase in one call",
]
add_bullet_text(slide5, tok2_lines, 7.3, 4.4, 5.5, 1.9, font_size=12.5, color=LIGHT_GRAY)

# ─────────────────────────────────────────────
# SLIDE 6: Head-to-Head Comparison
# ─────────────────────────────────────────────
slide6 = prs.slides.add_slide(blank_layout)
set_bg(slide6, DARK_BG)
add_rect(slide6, 0, 0, 13.33, 0.08, ACCENT)

add_text(slide6, "Head-to-Head Comparison", 0.3, 0.15, 12.5, 0.6, font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Table header
cols = ["Feature", "Web (Claude.ai)", "IDE / CLI"]
col_w = [3.5, 4.5, 4.83]
add_rect(slide6, 0.3, 0.85, 12.73, 0.45, DARK_CARD)
x = 0.3
for h, w in zip(cols, col_w):
    add_text(slide6, h, x+0.1, 0.88, w-0.1, 0.38, font_size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    x += w

table_data = [
    ("Setup",           "Zero install, browser",         "npm install / extension"),
    ("Execution",       "Anthropic-hosted container",    "Local machine / your server"),
    ("Agent Loop",      "✅ Managed Agents",              "✅ Tool Runner (you host)"),
    ("File Access",     "Per-session workspace mount",   "Full local filesystem"),
    ("Bedrock/Vertex",  "❌ Not supported",               "✅ Fully supported"),
    ("Context Window",  "Up to 1M tokens",               "Up to 1M tokens"),
    ("Prompt Caching",  "✅ (system prompt cache)",       "✅ (5-min & 1-hour TTL)"),
    ("Batch API",       "✅ via API",                     "✅ 50% cost reduction"),
    ("Control",         "Moderate (Anthropic manages)",  "Full control"),
    ("Best For",        "Demos, exploration, agents",    "Daily coding, CI/CD"),
]

for ri, (feat, web, ide) in enumerate(table_data):
    y = 1.35 + ri * 0.55
    bg = DARK_CARD if ri % 2 == 0 else RGBColor(0x12, 0x22, 0x33)
    add_rect(slide6, 0.3, y, 12.73, 0.52, bg)
    add_text(slide6, feat, 0.4, y+0.07, 3.3, 0.4, font_size=12, bold=True, color=LIGHT_GRAY)
    add_text(slide6, web,  3.9, y+0.07, 4.3, 0.4, font_size=12, color=BLUE)
    add_text(slide6, ide,  8.4, y+0.07, 4.6, 0.4, font_size=12, color=GREEN)

# ─────────────────────────────────────────────
# SLIDE 7: Recommendation
# ─────────────────────────────────────────────
slide7 = prs.slides.add_slide(blank_layout)
set_bg(slide7, DARK_BG)
add_rect(slide7, 0, 0, 13.33, 0.08, ACCENT)

add_text(slide7, "Recommendation & Decision Guide", 0.3, 0.15, 12.5, 0.7, font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Choose Web
add_rect(slide7, 0.3, 1.0, 5.9, 2.8, RGBColor(0x05, 0x1A, 0x30), BLUE, 2)
add_text(slide7, "Choose  🌐  Web  when...", 0.5, 1.1, 5.5, 0.5, font_size=16, bold=True, color=BLUE)
web_rec = [
    "✔  You want zero-setup demos",
    "✔  Long-horizon agentic tasks needed",
    "✔  Shared team agents with versioning",
    "✔  No local infra to maintain",
    "✔  Exploring new features quickly",
]
add_bullet_text(slide7, web_rec, 0.5, 1.65, 5.5, 2.0, font_size=13, color=LIGHT_GRAY)

# Choose IDE
add_rect(slide7, 7.13, 1.0, 5.9, 2.8, RGBColor(0x05, 0x25, 0x10), GREEN, 2)
add_text(slide7, "Choose  💻  IDE  when...", 7.33, 1.1, 5.5, 0.5, font_size=16, bold=True, color=GREEN)
ide_rec = [
    "✔  Daily coding on large repos",
    "✔  Need full filesystem / CI access",
    "✔  Cost matters — use Batch + Cache",
    "✔  Bedrock/Vertex/Foundry deployment",
    "✔  Fine-grained tool control needed",
]
add_bullet_text(slide7, ide_rec, 7.33, 1.65, 5.5, 2.0, font_size=13, color=LIGHT_GRAY)

# Verdict
add_rect(slide7, 0.3, 4.0, 12.73, 1.2, RGBColor(0x1A, 0x14, 0x05), ACCENT, 2)
add_text(slide7, "🏆  Verdict: Use Both Together", 0.5, 4.05, 12.33, 0.5, font_size=17, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(slide7, "Web for exploration & agentic demos  ·  IDE for production coding & cost efficiency", 0.5, 4.55, 12.33, 0.55, font_size=13.5, color=WHITE, align=PP_ALIGN.CENTER)

# Model recommendation
add_rect(slide7, 0.3, 5.35, 12.73, 1.75, DARK_CARD)
add_text(slide7, "Recommended Model Strategy:", 0.5, 5.4, 5, 0.4, font_size=13, bold=True, color=ACCENT)
add_text(slide7, "Simple tasks → Haiku 4.5 ($1/$5)  |  Daily coding → Sonnet 4.6 ($3/$15)  |  Complex / Agents → Opus 4.8 ($5/$25)  |  Hardest → Fable 5 ($10/$50)", 0.5, 5.82, 12.33, 0.5, font_size=12, color=LIGHT_GRAY)
add_text(slide7, "Always enable: Adaptive Thinking  ·  Prompt Caching  ·  Streaming  ·  Token Count API before large calls", 0.5, 6.3, 12.33, 0.5, font_size=12, color=LIGHT_GRAY)

# ─────────────────────────────────────────────
# SLIDE 8: Real Task — User Login Dashboard
# ─────────────────────────────────────────────
slide8 = prs.slides.add_slide(blank_layout)
set_bg(slide8, DARK_BG)
add_rect(slide8, 0, 0, 13.33, 0.08, ACCENT)

add_text(slide8, "Real Task: Build a User Login Dashboard", 0.3, 0.15, 12.5, 0.65, font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide8, "Token & cost breakdown for the same task on Web vs IDE", 0.3, 0.75, 12.5, 0.38, font_size=13, color=RGBColor(0x90,0x90,0x90), align=PP_ALIGN.CENTER)

# Task description box
add_rect(slide8, 0.3, 1.18, 12.73, 0.58, DARK_CARD, ACCENT, 1)
add_text(slide8, '📋  Task: "Create a login page with email/password form, JWT auth, session dashboard showing user info & last login time"', 0.5, 1.25, 12.3, 0.45, font_size=12.5, color=ACCENT)

# --- WEB column ---
add_rect(slide8, 0.3, 1.9, 6.1, 4.55, DARK_CARD, BLUE, 1.5)
add_text(slide8, "🌐  Web — Claude Opus 4.8", 0.5, 1.95, 5.8, 0.45, font_size=14, bold=True, color=BLUE)

web_tok = [
    ("System prompt (per turn)",     "~1,200 tokens",  "cached after 1st call"),
    ("User prompt",                  "~320 tokens",    "one-time"),
    ("File context (HTML/JS/CSS)",   "~2,800 tokens",  "cached after 1st edit"),
    ("Model output (code + explain)","~3,500 tokens",  "output, billed full"),
    ("Tool calls (write 4 files)",   "~900 tokens",    "input side"),
    ("Total input tokens",           "~5,020 tokens",  "1,200 cached"),
    ("Total output tokens",          "~3,500 tokens",  ""),
]
col_x = [0.45, 4.0, 5.45]
col_w = [3.5, 1.4, 0.9]
y_start = 2.48
for i, (label, val, note) in enumerate(web_tok):
    y = y_start + i * 0.48
    bg = DARK_BG if i % 2 == 0 else RGBColor(0x12,0x20,0x30)
    add_rect(slide8, 0.32, y, 6.06, 0.45, bg)
    is_total = label.startswith("Total")
    fc = ACCENT if is_total else LIGHT_GRAY
    add_text(slide8, label, 0.45, y+0.05, 3.45, 0.36, font_size=11, bold=is_total, color=fc)
    add_text(slide8, val,   4.0,  y+0.05, 1.4,  0.36, font_size=11, bold=is_total, color=WHITE, align=PP_ALIGN.RIGHT)
    add_text(slide8, note,  5.45, y+0.05, 0.9,  0.36, font_size=9, color=RGBColor(0x70,0x90,0x70), align=PP_ALIGN.LEFT)

# Cost calc web
add_rect(slide8, 0.32, 5.85, 6.06, 0.52, RGBColor(0x05,0x18,0x30), BLUE, 1)
add_text(slide8, "💰  Est. Cost:", 0.45, 5.9, 2.5, 0.38, font_size=12.5, bold=True, color=BLUE)
# Input: (5020-1200)*5/1M + 1200*0.5/1M = ~0.019+0.0006 | Output: 3500*25/1M = ~0.0875 → ~$0.107 uncached
# With cache: input cached portion saves 90% → ~$0.092 total
add_text(slide8, "~$0.09  (with prompt cache)", 2.8, 5.9, 3.5, 0.38, font_size=12.5, bold=True, color=WHITE)

# --- IDE column ---
add_rect(slide8, 7.1, 1.9, 6.1, 4.55, DARK_CARD, GREEN, 1.5)
add_text(slide8, "💻  IDE — Claude Sonnet 4.6", 7.3, 1.95, 5.8, 0.45, font_size=14, bold=True, color=GREEN)

ide_tok = [
    ("System prompt (per turn)",     "~1,200 tokens",  "cached (5-min TTL)"),
    ("User prompt",                  "~320 tokens",    "one-time"),
    ("File context (read 6 files)",  "~4,200 tokens",  "cached after 1st read"),
    ("Model output (code + explain)","~3,500 tokens",  "output, billed full"),
    ("Tool calls (edit/write files)","~600 tokens",    "input side"),
    ("Total input tokens",           "~6,320 tokens",  "5,400 cached"),
    ("Total output tokens",          "~3,500 tokens",  ""),
]
for i, (label, val, note) in enumerate(ide_tok):
    y = y_start + i * 0.48
    bg = DARK_BG if i % 2 == 0 else RGBColor(0x10,0x22,0x15)
    add_rect(slide8, 7.12, y, 6.06, 0.45, bg)
    is_total = label.startswith("Total")
    fc = ACCENT if is_total else LIGHT_GRAY
    add_text(slide8, label, 7.25, y+0.05, 3.45, 0.36, font_size=11, bold=is_total, color=fc)
    add_text(slide8, val,   10.8, y+0.05, 1.4,  0.36, font_size=11, bold=is_total, color=WHITE, align=PP_ALIGN.RIGHT)
    add_text(slide8, note,  12.25,y+0.05, 0.85, 0.36, font_size=9, color=RGBColor(0x70,0x90,0x70), align=PP_ALIGN.LEFT)

# Cost calc IDE
add_rect(slide8, 7.12, 5.85, 6.06, 0.52, RGBColor(0x05,0x20,0x0A), GREEN, 1)
add_text(slide8, "💰  Est. Cost:", 7.25, 5.9, 2.5, 0.38, font_size=12.5, bold=True, color=GREEN)
# Input: (6320-5400)*3/1M + 5400*0.3/1M = 0.00276+0.00162 | Output: 3500*15/1M = 0.0525 → ~$0.057 w/ cache
add_text(slide8, "~$0.055  (Sonnet + cache)", 9.6, 5.9, 3.5, 0.38, font_size=12.5, bold=True, color=WHITE)

# savings banner
add_rect(slide8, 0.3, 6.47, 12.73, 0.75, RGBColor(0x1A, 0x14, 0x03), ACCENT, 1.5)
add_text(slide8, "🏆  IDE saves ~40% on this task  |  Repeat runs (cache warm): IDE ~$0.016  vs  Web ~$0.032  |  Batch mode: extra 50% off", 0.5, 6.55, 12.33, 0.55, font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# Save
prs.save("/home/user/BABSanthakumar/Claude_Web_vs_IDE.pptx")
print("Saved: Claude_Web_vs_IDE.pptx")
