import tkinter as tk
from tkinter import filedialog
import threading
import traceback
import time
import os
import re
import math
from collections import Counter

try:
    from pypdf import PdfReader
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    import docx as _docx_lib
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

BG         = "#0D1B2A"
SIDEBAR_BG = "#091422"
HEADER_BG  = "#112240"
PANEL      = "#112240"
INPUT_FRM  = "#0A1A2E"
INPUT_BG   = "#172A45"
BORDER     = "#1E3A5F"
DIVIDER    = "#1A3050"
BTN_BG     = "#1A3A6E"

BLUE       = "#4A9EFF"
BLUE_DIM   = "#2E7DD6"

GOLD       = "#FFBC42"

TEXT_WHITE = "#F0F6FF"
TEXT_LIGHT = "#B8D4F0"
TEXT_MUTED = "#6B90B8"
TEXT_BLUE  = "#7EC8FF"

USER_BG    = "#0F2847"
USER_FG    = "#E8F4FF"
USER_LABEL = "#FFBC42"

BOT_BG     = "#0A1E35"
BOT_FG     = "#DCF0FF"
BOT_LABEL  = "#4A9EFF"

SRC_BG     = "#071428"
SRC_FG     = "#7EC8FF"

SUCCESS    = "#34D399"
WARN       = "#FBBF24"
ERROR      = "#F87171"
THINKING   = "#93C5FD"

FONT_MONO  = "Consolas"
FONT_SANS  = "Segoe UI"

PLACEHOLDER = "Ask a question about your document…"
MAX_FILE_MB = 20
MAX_HISTORY = 40

STOPWORDS = {
    "a","an","the","is","are","was","were","be","been","being","have","has",
    "had","do","does","did","will","would","could","should","may","might",
    "shall","can","need","dare","ought","used","to","of","in","for","on",
    "with","at","by","from","up","about","into","through","during","before",
    "after","above","below","between","out","off","over","under","then","once",
    "and","but","or","so","yet","both","either","neither","not","no","nor",
    "just","than","too","very","also","that","this","these","those","it","its",
    "i","you","he","she","we","they","them","their","what","which","who","whom",
    "how","when","where","why","all","each","every","few","more","most","other",
    "some","such","same","as","if","while","our","your","my"
}

def _tokenize(text: str) -> list:
    return re.findall(r'\b[a-zA-Z]{2,}\b', text.lower())


def _clean(tokens: list) -> list:
    return [t for t in tokens if t not in STOPWORDS]


_ABBREV_RE = re.compile(
    r'\b(Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|vs|etc|approx|Fig|Vol|Ed|no)\.',
    re.IGNORECASE
)
_ABBREV_PLACEHOLDER = "\x00"


def _sentences(text: str) -> list:

    text = re.sub(r'-\n', '', text)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r'([.!?])([A-Z])', r'\1 \2', text)

    text = _ABBREV_RE.sub(lambda m: m.group(0).replace('.', _ABBREV_PLACEHOLDER), text)

    raw  = re.split(r'(?<=[.!?])\s+', text)

    return [s.replace(_ABBREV_PLACEHOLDER, '.').strip()
            for s in raw if len(s.strip()) > 20]


def _jaccard(a: str, b: str) -> float:
    sa, sb = set(_tokenize(a)), set(_tokenize(b))
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def _tfidf_from_tokens(qtoks: list, stoks_counter: Counter,
                       stoks_len: int, idf_cache: dict) -> float:

    if stoks_len == 0:
        return 0.0
    score = 0.0
    for qt in qtoks:
        tf  = stoks_counter.get(qt, 0) / stoks_len
        idf = idf_cache.get(qt, 1.0)
        score += tf * idf
    return score


def _overlap_from_tokens(qtoks: list, stoks_set: set) -> float:

    if not stoks_set:
        return 0.0
    return sum(1 for qt in qtoks if qt in stoks_set) / max(len(qtoks), 1)


def _qtype(q: str) -> str:

    q = q.lower().strip()

    if q.startswith(("how many", "how much", "how long", "how often", "how far")):
        return "quantity"

    if q.startswith(("what year", "what date", "what time", "when was",
                     "when did", "when is", "when are", "when were")):
        return "time"

    if q.startswith(("where", "which place", "which location",
                     "which country", "which city")):
        return "location"

    if q.startswith(("who ", "whose ", "whom ")):
        return "person"

    if q.startswith(("why ", "what reason", "what cause")):
        return "reason"

    if q.startswith(("how ", "in what way", "by what")):
        return "method"

    if q.startswith(("what is", "what are", "what was", "what were",
                     "what does", "what do")):
        return "definition"

    if q.split()[0] in {"is","are","was","were","does","do","did","has",
                        "have","can","could","would","should","will","won't",
                        "isn't","aren't","doesn't","don't","didn't"}:
        return "yesno"

    if q.startswith("what"):
        return "definition"
    return "general"


PREFIXES = {
    "definition": "According to the document, ",
    "quantity":   "The document indicates that ",
    "time":       "The document states that ",
    "location":   "As stated in the document, ",
    "person":     "The document refers to ",
    "reason":     "The document explains that ",
    "method":     "According to the document, ",
    "general":    "Based on the document, ",
    "yesno":      "Based on the document, ",
}


def _synthesize(top_sents: list, qtype: str, max_score: float) -> tuple:

    if max_score >= 0.25:
        confidence = "high"
    elif max_score >= 0.10:
        confidence = "medium"
    else:
        confidence = "low"

    if not top_sents:
        return (
            "I couldn't find relevant information in this document to answer that.\n"
            "Try rephrasing, or the topic may not be covered.",
            "none"
        )

    if qtype == "yesno":
        evidence = " ".join(top_sents[:2])
        answer   = f"Based on the document:\n\n  {evidence}"
        return answer, confidence

    # General answer
    combined = " ".join(top_sents[:3])

    if len(combined) > 520:
        trimmed   = combined[:520]
        positions = [trimmed.rfind(p) for p in ".!?" if trimmed.rfind(p) > 200]
        if positions:
            combined = trimmed[:max(positions) + 1]

    prefix = PREFIXES.get(qtype, "Based on the document, ")
    if combined and combined[0].isupper():
        combined = combined[0].lower() + combined[1:]
    return prefix + combined, confidence

class LocalAIEngine:

    def __init__(self):
        self.sentences  : list  = []
        self.idf_cache  : dict  = {}
        self._tok_cache : list  = []

    def index(self, text: str):
        self.sentences = _sentences(text)
        self._build_idf_cache()

    def _build_idf_cache(self):

        N = len(self.sentences)
        if N == 0:
            self.idf_cache  = {}
            self._tok_cache = []
            return
        self._tok_cache = []
        df_counts: dict = {}
        for s in self.sentences:
            toks  = _clean(_tokenize(s))
            ctr   = Counter(toks)
            s_set = set(toks)
            self._tok_cache.append((ctr, s_set, len(toks)))

            for word in s_set:
                df_counts[word] = df_counts.get(word, 0) + 1


        self.idf_cache = {
            word: math.log((N + 1) / (df + 1)) + 1
            for word, df in df_counts.items()
        }

    def answer(self, question: str) -> tuple:

        if not self.sentences:
            return "No document has been loaded yet.", "none", []

        qt    = _qtype(question)
        qtoks = _clean(_tokenize(question))
        if not qtoks:
            return "Please ask a more specific question.", "none", []

        total = len(self.sentences)
        scored = []
        for i, (s, (ctr, s_set, tlen)) in enumerate(
                zip(self.sentences, self._tok_cache)):

            tf_score  = _tfidf_from_tokens(qtoks, ctr, tlen, self.idf_cache)
            ov_score  = _overlap_from_tokens(qtoks, s_set)
            raw_score = 0.55 * tf_score + 0.45 * ov_score

            if i < total * 0.10 or i > total * 0.90:
                raw_score *= 1.20

            scored.append((raw_score, s))

        scored.sort(reverse=True)
        max_score = scored[0][0] if scored else 0.0

        # Threshold + dedup
        top = [s for sc, s in scored[:8] if sc >= 0.04]
        deduped = []
        for s in top:
            if not any(_jaccard(s, k) > 0.72 for k in deduped):
                deduped.append(s)

        answer_text, confidence = _synthesize(deduped[:3], qt, max_score)
        return answer_text, confidence, deduped[:2]

class TypingAnimator:

    def __init__(self, widget, text, tag, on_done=None):
        self.widget  = widget
        self.text    = text
        self.tag     = tag
        self.on_done = on_done
        self.idx     = 0
        self._job    = None

        length     = max(len(text), 1)
        self.delay = max(4, min(14, 1200 // length))

    def start(self):
        self._tick()

    def _tick(self):
        if self.idx < len(self.text):
            self.widget.config(state="normal")
            self.widget.insert(tk.END, self.text[self.idx], self.tag)
            self.widget.config(state="disabled")
            self.widget.see(tk.END)
            self.idx += 1
            self._job = self.widget.after(self.delay, self._tick)
        else:
            if self.on_done:
                self.on_done()

    def cancel(self):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None

class AuraChat:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("AuraChat  ·  Document Intelligence")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h   = min(1260, int(sw * 0.87)), min(860, int(sh * 0.87))
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.minsize(820, 600)

        try:
            icon = tk.PhotoImage(width=32, height=32)
            for y in range(32):
                for x in range(32):
                    icon.put("#0D1B2A", (x, y))
            icon.put("#4A9EFF", (8, 28));  icon.put("#4A9EFF", (9, 28))
            icon.put("#4A9EFF", (10,24)); icon.put("#4A9EFF", (11,24))
            icon.put("#4A9EFF", (12,20)); icon.put("#4A9EFF", (13,20))
            icon.put("#4A9EFF", (14,16)); icon.put("#4A9EFF", (15,16))
            icon.put("#4A9EFF", (16,16)); icon.put("#4A9EFF", (17,20))
            icon.put("#4A9EFF", (18,20)); icon.put("#4A9EFF", (19,24))
            icon.put("#4A9EFF", (20,24)); icon.put("#4A9EFF", (21,28))
            icon.put("#4A9EFF", (22,28)); icon.put("#4A9EFF", (23,28))
            icon.put("#4A9EFF", (12,22)); icon.put("#4A9EFF", (13,22))
            icon.put("#4A9EFF", (14,22)); icon.put("#4A9EFF", (15,22))
            icon.put("#4A9EFF", (16,22)); icon.put("#4A9EFF", (17,22))
            icon.put("#4A9EFF", (18,22)); icon.put("#4A9EFF", (19,22))
            root.iconphoto(True, icon)
        except Exception:
            pass

        self.ai           = LocalAIEngine()
        self.doc_loaded   = False
        self._is_indexing = False
        self._msg_count   = 0
        self._anim        = None
        self._dots_job    = None
        self._dot_mark    = None
        self._dots_count  = 0
        self._ph_active   = True
        self._ew_ref      = None

        self._build()

        # FIX #12 — keyboard shortcuts
        root.bind("<Control-o>", lambda _: self._upload())
        root.bind("<Control-O>", lambda _: self._upload())
        root.bind("<Control-l>", lambda _: self._clear())
        root.bind("<Control-L>", lambda _: self._clear())
        root.bind("<Control-e>", lambda _: self._export())
        root.bind("<Control-E>", lambda _: self._export())
        root.bind("<Escape>",    lambda _: self._cancel_anim())

    def _build(self):
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()
        self._greet()

    def _build_sidebar(self):
        sb = tk.Frame(self.root, bg=SIDEBAR_BG, width=258)
        sb.grid(row=0, column=0, sticky="ns")
        sb.grid_propagate(False)
        sb.columnconfigure(0, weight=1)
        sb.rowconfigure(8, weight=1)

        # Logo
        logo = tk.Frame(sb, bg=SIDEBAR_BG)
        logo.grid(row=0, column=0, padx=22, pady=(38, 2), sticky="w")
        tk.Label(logo, text="AURA", font=(FONT_MONO, 22, "bold"),
                 bg=SIDEBAR_BG, fg=BLUE).pack(side="left")
        tk.Label(logo, text="CHAT", font=(FONT_MONO, 22, "bold"),
                 bg=SIDEBAR_BG, fg=TEXT_WHITE).pack(side="left")

        tk.Label(sb, text="DOCUMENT INTELLIGENCE",
                 font=(FONT_MONO, 7, "bold"),
                 bg=SIDEBAR_BG, fg=TEXT_MUTED
                 ).grid(row=1, column=0, padx=22, pady=(0, 28), sticky="w")

        tk.Frame(sb, bg=DIVIDER, height=1
                 ).grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 18))

        self._side_btn(sb, row=3, icon="📂", label="UPLOAD   Ctrl+O", cmd=self._upload)
        self._side_btn(sb, row=4, icon="💾", label="EXPORT   Ctrl+E", cmd=self._export)
        self._side_btn(sb, row=5, icon="🗑", label="CLEAR    Ctrl+L", cmd=self._clear)

        tk.Frame(sb, bg=SIDEBAR_BG).grid(row=8, column=0, sticky="nsew")

        tk.Frame(sb, bg=DIVIDER, height=1
                 ).grid(row=9, column=0, sticky="ew", padx=16, pady=(0, 10))

        badge = tk.Frame(sb, bg=SIDEBAR_BG)
        badge.grid(row=10, column=0, padx=20, pady=(0, 8), sticky="w")
        tk.Label(badge, text="⚡", font=(FONT_SANS, 12),
                 bg=SIDEBAR_BG, fg=GOLD).pack(side="left")
        inf = tk.Frame(badge, bg=SIDEBAR_BG)
        inf.pack(side="left", padx=(8, 0))
        tk.Label(inf, text="BUILT-IN AI ENGINE",
                 font=(FONT_MONO, 8, "bold"),
                 bg=SIDEBAR_BG, fg=GOLD).pack(anchor="w")
        tk.Label(inf, text="No API · No Internet · No Download",
                 font=(FONT_MONO, 7),
                 bg=SIDEBAR_BG, fg=TEXT_MUTED).pack(anchor="w")

        self.doc_status = tk.Label(
            sb, text="No document loaded",
            font=(FONT_MONO, 7), bg=SIDEBAR_BG,
            fg=TEXT_MUTED, wraplength=218, justify="left")
        self.doc_status.grid(row=11, column=0, padx=20, pady=(4, 22), sticky="w")

    def _side_btn(self, parent, row, icon, label, cmd):
        frm = tk.Frame(parent, bg=SIDEBAR_BG)
        frm.grid(row=row, column=0, padx=14, pady=3, sticky="ew")
        btn = tk.Button(frm, text=f"  {icon}  {label}", command=cmd,
                        bg=BTN_BG, fg=TEXT_WHITE,
                        font=(FONT_MONO, 8, "bold"),
                        relief="flat", pady=10, padx=12,
                        cursor="hand2", anchor="w",
                        activebackground=BORDER, activeforeground=BLUE, bd=0)
        btn.pack(fill="x")
        btn.bind("<Enter>", lambda e, b=btn: b.config(bg=BORDER, fg=BLUE))
        btn.bind("<Leave>", lambda e, b=btn: b.config(bg=BTN_BG,  fg=TEXT_WHITE))

    def _build_main(self):
        main = tk.Frame(self.root, bg=BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(2, weight=1)
        main.columnconfigure(0, weight=1)

        # Header
        hdr = tk.Frame(main, bg=HEADER_BG, height=50)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)
        hdr.grid_propagate(False)

        tk.Label(hdr, text="Document Q&A",
                 font=(FONT_MONO, 13, "bold"),
                 bg=HEADER_BG, fg=TEXT_WHITE
                 ).grid(row=0, column=0, padx=26, pady=13, sticky="w")

        self.hdr_doc = tk.Label(hdr, text="No document loaded",
                                font=(FONT_MONO, 8),
                                bg=HEADER_BG, fg=TEXT_MUTED)
        self.hdr_doc.grid(row=0, column=2, padx=22, pady=13, sticky="e")

        tk.Frame(main, bg=BLUE, height=2).grid(row=1, column=0, sticky="ew")

        # Chat
        chat_wrap = tk.Frame(main, bg=BG)
        chat_wrap.grid(row=2, column=0, sticky="nsew")
        chat_wrap.rowconfigure(0, weight=1)
        chat_wrap.columnconfigure(0, weight=1)

        self.chat = tk.Text(
            chat_wrap, bg=BG, fg=TEXT_WHITE,
            font=(FONT_SANS, 11),
            borderwidth=0, highlightthickness=0,
            padx=26, pady=18, wrap="word", cursor="arrow",
            insertbackground=BLUE,
            selectbackground=BLUE_DIM, selectforeground=TEXT_WHITE,
            spacing3=2,
        )
        self.chat.grid(row=0, column=0, sticky="nsew")

        vsb = tk.Scrollbar(chat_wrap, orient="vertical",
                           command=self.chat.yview,
                           bg=HEADER_BG, troughcolor=BG,
                           highlightthickness=0, bd=0, width=7)
        vsb.grid(row=0, column=1, sticky="ns")
        self.chat.config(yscrollcommand=vsb.set)

        # Tags
        self.chat.tag_configure("you_label",
            foreground=USER_LABEL, font=(FONT_MONO, 8, "bold"), spacing1=22)
        self.chat.tag_configure("you_msg",
            background=USER_BG, foreground=USER_FG, font=(FONT_SANS, 11),
            lmargin1=55, lmargin2=55, rmargin=18, spacing1=5, spacing3=14)

        self.chat.tag_configure("bot_label",
            foreground=BOT_LABEL, font=(FONT_MONO, 8, "bold"), spacing1=22)
        self.chat.tag_configure("bot_msg",
            background=BOT_BG, foreground=BOT_FG, font=(FONT_SANS, 11),
            lmargin1=18, lmargin2=18, rmargin=55, spacing1=5, spacing3=14)

        self.chat.tag_configure("src_label",
            foreground=TEXT_MUTED, font=(FONT_MONO, 7, "bold"), spacing1=4)
        self.chat.tag_configure("src_msg",
            background=SRC_BG, foreground=SRC_FG,
            font=(FONT_SANS, 9, "italic"),
            lmargin1=18, lmargin2=30, rmargin=55, spacing1=2, spacing3=10)

        self.chat.tag_configure("conf_high",
            foreground=SUCCESS, font=(FONT_MONO, 7, "bold"),
            lmargin1=18, spacing1=2, spacing3=2)
        self.chat.tag_configure("conf_med",
            foreground=WARN, font=(FONT_MONO, 7, "bold"),
            lmargin1=18, spacing1=2, spacing3=2)
        self.chat.tag_configure("conf_low",
            foreground=ERROR, font=(FONT_MONO, 7, "bold"),
            lmargin1=18, spacing1=2, spacing3=2)

        self.chat.tag_configure("thinking",
            foreground=THINKING, font=(FONT_SANS, 11, "italic"),
            lmargin1=18, lmargin2=18, spacing1=5, spacing3=14)

        self.chat.tag_configure("error_label",
            foreground=ERROR, font=(FONT_MONO, 8, "bold"), spacing1=22)
        self.chat.tag_configure("error_msg",
            background="#1F0A0A", foreground="#FFB3B3", font=(FONT_SANS, 11),
            lmargin1=18, lmargin2=18, rmargin=55, spacing1=5, spacing3=14)

        self.chat.tag_configure("info_line",
            foreground=TEXT_BLUE, font=(FONT_SANS, 9, "italic"),
            lmargin1=18, spacing1=6, spacing3=2)

        self.chat.config(state="disabled")

        # Status bar
        sbar = tk.Frame(main, bg=HEADER_BG, height=32)
        sbar.grid(row=3, column=0, sticky="ew")
        sbar.grid_propagate(False)
        sbar.columnconfigure(1, weight=1)

        self.s_dot = tk.Label(sbar, text="●",
                              font=(FONT_SANS, 9), bg=HEADER_BG, fg=SUCCESS)
        self.s_dot.grid(row=0, column=0, padx=(18, 5), pady=7)
        self.s_txt = tk.Label(sbar, text="System Ready  |  Ctrl+O Upload  Ctrl+E Export  Ctrl+L Clear",
                              font=(FONT_MONO, 7), bg=HEADER_BG, fg=TEXT_LIGHT)
        self.s_txt.grid(row=0, column=1, sticky="w", pady=7)

        inp_frame = tk.Frame(main, bg=INPUT_FRM, pady=12, padx=18)
        inp_frame.grid(row=4, column=0, sticky="ew")
        inp_frame.columnconfigure(0, weight=1)

        ew = tk.Frame(inp_frame, bg=INPUT_BG,
                      highlightbackground=BORDER,
                      highlightcolor=BLUE,
                      highlightthickness=2)
        ew.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        ew.columnconfigure(0, weight=1)
        self._ew_ref = ew

        self.entry = tk.Entry(
            ew, font=(FONT_SANS, 12),
            bg=INPUT_BG, fg=TEXT_MUTED,
            insertbackground=BLUE,
            relief="flat", highlightthickness=0, bd=0,
        )
        self.entry.grid(row=0, column=0, sticky="ew", ipady=12, padx=12, pady=2)
        self.entry.insert(0, PLACEHOLDER)

        # FIX #3 — each event bound exactly ONCE
        self.entry.bind("<FocusIn>",  self._on_focus_in)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<Return>",   lambda _: self._ask())

        self.ask_btn = tk.Button(
            inp_frame, text="ASK  ›", command=self._ask,
            bg=BLUE, fg="#FFFFFF",
            font=(FONT_MONO, 10, "bold"),
            padx=28, pady=12, relief="flat", cursor="hand2",
            activebackground=BLUE_DIM, activeforeground="#FFFFFF", bd=0,
        )
        self.ask_btn.grid(row=0, column=1)
        self.ask_btn.bind("<Enter>", lambda _: self.ask_btn.config(bg=BLUE_DIM))
        self.ask_btn.bind("<Leave>", lambda _: self.ask_btn.config(bg=BLUE))

    def _on_focus_in(self, _=None):
        self._ew_ref.config(highlightbackground=BLUE)
        if self._ph_active:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=TEXT_WHITE)
            self._ph_active = False

    def _on_focus_out(self, _=None):
        self._ew_ref.config(highlightbackground=BORDER)
        if not self.entry.get().strip():
            self.entry.insert(0, PLACEHOLDER)
            self.entry.config(fg=TEXT_MUTED)
            self._ph_active = True

    def _status(self, msg: str, kind: str = "ok"):
        col = {"ok": SUCCESS, "warn": WARN, "error": ERROR, "think": THINKING}
        self.s_dot.config(fg=col.get(kind, TEXT_MUTED))
        self.s_txt.config(text=msg)
        self.root.update_idletasks()

    def _start_thinking(self):
        self.chat.config(state="normal")
        self.chat.insert(tk.END, "  AURACHAT\n", "bot_label")
        self._dot_mark   = self.chat.index(tk.END)
        self._dots_count = 0
        self.chat.insert(tk.END, "  Thinking  ●\n", "thinking")
        self.chat.config(state="disabled")
        self.chat.see(tk.END)
        self._animate_dots()

    def _animate_dots(self):

        if not self._dot_mark:
            return
        frames = ["  Thinking  ●", "  Thinking  ● ●",
                  "  Thinking  ● ● ●", "  Thinking  ● ●"]
        frame = frames[self._dots_count % len(frames)]
        self._dots_count += 1
        try:
            self.chat.config(state="normal")
            s = self._dot_mark
            e = self.chat.index(f"{s} lineend +1c")
            self.chat.delete(s, e)
            self.chat.insert(s, f"{frame}\n", "thinking")
            self.chat.config(state="disabled")
            self.chat.see(tk.END)
        except tk.TclError:

            self._dot_mark = None
            return
        self._dots_job = self.root.after(380, self._animate_dots)

    def _stop_thinking_type(self, answer: str, sources: list,
                            confidence: str, on_done=None):
        if self._dots_job:
            self.root.after_cancel(self._dots_job)
            self._dots_job = None
        # Remove thinking placeholder
        self.chat.config(state="normal")
        if self._dot_mark:
            try:
                s = self._dot_mark
                e = self.chat.index(f"{s} lineend +1c")
                self.chat.delete(s, e)
            except tk.TclError:
                pass
        self._dot_mark = None
        self.chat.config(state="disabled")

        # Stream the answer
        indented = "\n".join(f"  {l}" for l in answer.splitlines()) + "\n"

        def _after_answer():
            if sources:
                self._w(("  📌 Sources from document:\n", "src_label"))
                for i, src in enumerate(sources, 1):
                    short = src[:130] + "…" if len(src) > 130 else src
                    self._w((f"  [{i}] {short}\n", "src_msg"))

            conf_tags = {"high": "conf_high", "medium": "conf_med",
                         "low": "conf_low",   "none": "conf_low"}
            conf_labels = {
                "high":   "  ● High confidence\n\n",
                "medium": "  ◐ Medium confidence — answer may be partial\n\n",
                "low":    "  ○ Low confidence — topic may not be well covered\n\n",
                "none":   "\n",
            }
            tag   = conf_tags.get(confidence, "conf_low")
            label = conf_labels.get(confidence, "\n")
            self._w((label, tag))
            if on_done:
                on_done()

        anim = TypingAnimator(self.chat, indented, "bot_msg",
                              on_done=_after_answer)
        self._anim = anim
        anim.start()

    def _upload(self):
        if self._is_indexing:
            return
        ftypes = [("Supported", "*.pdf *.txt"), ("PDF", "*.pdf"),
                  ("Text", "*.txt")]
        if DOCX_SUPPORT:
            ftypes.insert(1, ("Word", "*.docx"))
            ftypes[0] = ("Supported", "*.pdf *.txt *.docx")

        path = filedialog.askopenfilename(title="Select a Document",
                                          filetypes=ftypes)
        if not path:
            return

        try:
            mb = os.path.getsize(path) / (1024 * 1024)
            if mb > MAX_FILE_MB:
                self._add_info(
                    f"⚠  Large file ({mb:.1f} MB) — indexing may take a moment…")
        except OSError:
            pass

        self.doc_loaded   = False
        self._is_indexing = True
        self.ask_btn.config(state="disabled", bg=BLUE_DIM)
        self._status("Indexing document…", "warn")

        def _run():
            try:
                ext  = os.path.splitext(path)[1].lower()
                text = ""

                if ext == ".txt":
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        text = f.read()

                elif ext == ".pdf":
                    if not PDF_SUPPORT:
                        raise RuntimeError(
                            "pypdf not installed.\nRun:  pip install pypdf")
                    reader = PdfReader(path)
                    text   = "\n".join(
                        p.extract_text() for p in reader.pages
                        if p.extract_text())

                elif ext == ".docx":                          # FIX #23
                    if not DOCX_SUPPORT:
                        raise RuntimeError(
                            "python-docx not installed.\n"
                            "Run:  pip install python-docx")
                    doc  = _docx_lib.Document(path)
                    text = "\n".join(
                        p.text for p in doc.paragraphs if p.text.strip())

                else:
                    raise ValueError(f"Unsupported file type: {ext}")

                if not text.strip():
                    raise ValueError("The document appears empty or unreadable.")

                self.ai.index(text)
                name  = os.path.basename(path)
                sents = len(self.ai.sentences)
                # FIX #2 — doc_loaded set on main thread only
                self.root.after(0, lambda: self._on_loaded(name, sents))

            except Exception as ex:
                # FIX #14 — print full traceback to console
                print(f"\n[AuraChat Upload Error]\n{traceback.format_exc()}")
                self.root.after(0, lambda: self._on_err(str(ex)))

        threading.Thread(target=_run, daemon=True).start()

    def _on_loaded(self, name: str, sents: int):

        self.doc_loaded   = True
        self._is_indexing = False
        self.ask_btn.config(state="normal", bg=BLUE)

        short = name[:34] + ("…" if len(name) > 34 else "")
        self._status(f"Indexed {sents} sentences — ready!", "ok")
        self.hdr_doc.config(
            text=f"📄  {short}  ·  {sents} sentences", fg=SUCCESS)
        self.doc_status.config(
            text=f"📄 {name}\n{sents} sentences indexed", fg=SUCCESS)
        self._add_info(f"✓  Loaded: {name}  ({sents} sentences indexed)")
        self._add_bot(
            f"Document loaded!\n\n"
            f"  File:     {name}\n"
            f"  Indexed:  {sents} sentences\n\n"
            f"Ready — ask me anything about this document."
        )

    def _on_err(self, msg: str):
        self._is_indexing = False
        self.ask_btn.config(state="normal", bg=BLUE)
        self._status("Error loading document", "error")
        self._add_error("Upload Failed", msg)
        self._show_err("Upload Error", msg)
    def _ask(self):
        raw = self.entry.get().strip()

        if not raw or raw == PLACEHOLDER:
            return
        if not self.doc_loaded:
            self._show_err(
                "No Document Loaded",
                "Please upload a document first.\n\n"
                "Click  📂 UPLOAD DOCUMENT  in the sidebar\n"
                "or press  Ctrl+O")
            return

        raw = raw[:500]

        self._cancel_anim()
        self._add_user(raw)
        self.entry.delete(0, tk.END)
        self._on_focus_out()
        self.ask_btn.config(state="disabled", bg=BLUE_DIM)
        self._status("Thinking…", "think")
        self._start_thinking()

        self._msg_count += 1
        if self._msg_count > MAX_HISTORY:
            self._trim_history()

        def _run():
            time.sleep(0.2)
            answer, confidence, sources = self.ai.answer(raw)
            self.root.after(
                0, lambda: self._deliver(answer, confidence, sources))

        threading.Thread(target=_run, daemon=True).start()

    def _deliver(self, answer: str, confidence: str, sources: list):
        def _done():
            self.chat.config(state="normal")
            self.chat.config(state="disabled")
            self.ask_btn.config(state="normal", bg=BLUE)
            self._status("Ready  |  Ctrl+O Upload  Ctrl+E Export  Ctrl+L Clear",
                         "ok")
        self._stop_thinking_type(answer, sources, confidence, on_done=_done)

    def _trim_history(self):

        self.chat.config(state="normal")
        try:

            idx = "1.0"
            for _ in range(10):
                idx = self.chat.index(f"{idx} +1 line")
            self.chat.delete("1.0", idx)
        except tk.TclError:
            pass
        self.chat.config(state="disabled")
        self._msg_count = MAX_HISTORY   # reset counter

    def _export(self):
        content = self.chat.get("1.0", tk.END).strip()
        if not content:
            self._show_err("Nothing to Export", "The chat history is empty.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Conversation",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt")],
            initialfile=f"aurachat_{time.strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(f"AuraChat Export  —  {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 60 + "\n\n")
                f.write(content)
            self._add_info(
                f"✓  Conversation exported to  {os.path.basename(path)}")
            self._status("Exported successfully", "ok")
        except Exception as ex:
            print(f"\n[AuraChat Export Error]\n{traceback.format_exc()}")
            self._show_err("Export Failed", str(ex))

    def _w(self, *parts):
        self.chat.config(state="normal")
        for text, tag in parts:
            self.chat.insert(tk.END, text, tag)
        self.chat.config(state="disabled")
        self.chat.see(tk.END)

    def _add_user(self, text: str):
        ind = "\n".join(f"  {l}" for l in text.splitlines()) + "\n"
        self._w(("  YOU\n", "you_label"), (ind, "you_msg"), ("\n", "you_msg"))

    def _add_bot(self, text: str):
        ind = "\n".join(f"  {l}" for l in text.splitlines()) + "\n"
        self._w(("  AURACHAT\n", "bot_label"), (ind, "bot_msg"), ("\n", "bot_msg"))

    def _add_error(self, lbl: str, text: str):
        ind = "\n".join(f"  {l}" for l in text.splitlines()) + "\n"
        self._w((f"  ⚠  {lbl}\n", "error_label"),
                (ind, "error_msg"), ("\n", "error_msg"))

    def _add_info(self, text: str):
        self._w((f"  {text}\n\n", "info_line"))

    def _greet(self):
        self._add_bot(
            "Welcome to AuraChat — Document Intelligence.\n\n"
            "  1.  Press  Ctrl+O  or click  📂 UPLOAD  in the sidebar\n"
            "  2.  Load any  .pdf  /  .txt  /  .docx  file\n"
            "  3.  Type your question and press  ASK ›  or  Enter\n\n"
            "  Shortcuts:  Ctrl+O Upload  ·  Ctrl+E Export  ·  Ctrl+L Clear\n\n"
            "Built-in AI engine — no internet, no API key, no model download."
        )

    def _cancel_anim(self):
        if self._anim:
            self._anim.cancel()
            self._anim = None

    def _clear(self):
        self._cancel_anim()
        if self._dots_job:
            self.root.after_cancel(self._dots_job)
            self._dots_job = None
        self._dot_mark  = None   # FIX #6 — reset mark so animator guard triggers
        self._msg_count = 0
        self.chat.config(state="normal")
        self.chat.delete("1.0", tk.END)
        self.chat.config(state="disabled")
        self._status("History cleared  |  Ctrl+O Upload  Ctrl+E Export", "ok")
        self._greet()

    def _show_err(self, title: str, msg: str):
        top = tk.Toplevel(self.root)
        top.title(title)
        top.configure(bg=PANEL)
        top.resizable(False, False)
        top.grab_set()
        pw, ph = 440, 215
        px = self.root.winfo_x() + (self.root.winfo_width()  - pw) // 2
        py = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        top.geometry(f"{pw}x{ph}+{px}+{py}")

        tk.Label(top, text=f"⚠  {title}",
                 font=(FONT_MONO, 11, "bold"),
                 bg=PANEL, fg=ERROR, pady=16).pack()
        tk.Label(top, text=msg,
                 font=(FONT_SANS, 10), bg=PANEL, fg=TEXT_WHITE,
                 wraplength=380, justify="left").pack(padx=28)
        tk.Button(top, text="  OK  ", command=top.destroy,
                  bg=BLUE, fg="#FFFFFF",
                  font=(FONT_MONO, 9,  "bold"),
                  relief="flat", padx=24, pady=8,
                  cursor="hand2").pack(pady=16)
        top.bind("<Return>", lambda _: top.destroy())
        top.bind("<Escape>", lambda _: top.destroy())

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.tk.call('tk', 'scaling', 1.25)
    except Exception:
        pass
    AuraChat(root)
    root.mainloop()