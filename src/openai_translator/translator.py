import os
import uuid
import json
import time
import tiktoken
import mysql.connector
from openai import OpenAI
import re

class OpenAITranslator:
    def __init__(self, model="gpt-5-nano", filename=None):
        self.model = model
        self.api_key = os.environ.get("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError("OPENAI_API_KEY not set in environment")
        self.client = OpenAI(api_key=self.api_key)

        # DB config from environment
        self.db_config = {
            "host": os.environ.get("MARIADB_HOST", "localhost"),
            "user": os.environ.get("MARIADB_USER", "root"),
            "password": os.environ.get("MARIADB_PASSWORD", ""),
            "database": os.environ.get("MARIADB_DB", "translation")
        }

        self.doc_id = None
        self.filename = None

        if filename:
            self.set_filename(filename)

    def set_filename(self, filename):
        """Change active document context."""
        self.filename = filename
        self.doc_id = str(uuid.uuid4())
        try:
            conn = self.get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO documents (doc_id, filename) VALUES (%s, %s)",
                (self.doc_id, self.filename)
            )
            conn.commit()
            print(f"[INFO] Created document: {self.filename} ({self.doc_id})")
        except Exception as e:
            print(f"[Warning] Failed to save document info: {e}")
        finally:
            try: cursor.close(); conn.close()
            except: pass

    def get_db_connection(self):
        return mysql.connector.connect(**self.db_config)

    @staticmethod
    def estimate_tokens(text: str) -> int:
        encoding = tiktoken.get_encoding("cl100k_base")
        return len(encoding.encode(text))

    @staticmethod
    def calculate_openai_cost(response_json):
        model = response_json.get("model", "")
        usage = response_json.get("usage", {})
        prompt_tokens = usage.get("prompt_tokens", 0)
        completion_tokens = usage.get("completion_tokens", 0)
        
        PRICES = {
            "gpt-5-pro": {"input": 15, "output": 120},
            "gpt-5.1": {"input": 1.25, "output": 10.00},  # <-- new line for gpt-5.1
            "gpt-5": {"input": 1.25, "output": 10.00},
            "gpt-5-mini": {"input": 0.25, "output": 2.00},
            "gpt-5-nano": {"input": 0.05, "output": 0.40},
            "gpt-4o": {"input": 2.50, "output": 10.00},
            "gpt-4o-mini": {"input": 0.15, "output": 0.60}
        }
        
        # Find matching model price tier (partial match supported)
        price = next((v for k, v in PRICES.items() if k in model), None)
        
        if price is None:
            print(f"[WARN] No known pricing for model '{model}'. Cost will be set to 0.")
            return {
                "model": model,
                "prompt_tokens": prompt_tokens,
                "completion_tokens": completion_tokens,
                "total_tokens": prompt_tokens + completion_tokens,
                "input_cost_usd": 0.0,
                "output_cost_usd": 0.0,
                "total_cost_usd": 0.0
            }
        
        input_cost = (prompt_tokens / 1_000_000) * price["input"]
        output_cost = (completion_tokens / 1_000_000) * price["output"]
        total_cost = input_cost + output_cost
        
        print("Total cost in USD:", round(total_cost, 6))
        
        return {
            "model": model,
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": prompt_tokens + completion_tokens,
            "input_cost_usd": round(input_cost, 6),
            "output_cost_usd": round(output_cost, 6),
            "total_cost_usd": round(total_cost, 6)
        }

    @staticmethod
    def build_translation_prompt(source_lang, dest_lang, text):
        lines = text.split("\n")
        numbered_lines = [f"Line {i+1}: {line}" for i, line in enumerate(lines)]
        numbered_text = "\n".join(numbered_lines)
        
        prompt = (
            f"You are a professional subtitling translator.\n"
            f"Your task is to translate {source_lang} into high-quality {dest_lang} suitable for television subtitles.\n"

            f"Overall context and style:\n"
            f"Read all lines first so you understand the full context, intent, and information hierarchy.\n"
            f"Treat the entire input as one coherent text when choosing tone, terminology, and phrasing.\n"
            f"This global understanding is only for lexical choice, tone, and consistency; "
            f"do NOT redistribute, move, or re-balance information across lines.\n"
            f"Ensure consistent translations for recurring terms, names, and concepts across all lines.\n"
            f"If part of the text to translate is already in {dest_lang}, treat it as authoritative translation memory and keep it literal.\n"

            f"Line-by-line constraints:\n"
            f"Translate line by line: produce exactly one output line for each input line.\n"
            f"Do NOT merge, split, add, remove, or repeat lines.\n"
            f"Preserve the input line order.\n"
            f"Parentheses and multiple sentences within a line belong to that same line.\n"

            f"Stylistic and linguistic rules:\n"
            f"Use a formal, standard, and natural register appropriate for broadcast media; "
            f"avoid colloquial speech and overly literary or archaic language, while preserving all core information.\n"
            f"The wording inside each line may become slightly shorter or longer to produce natural {dest_lang}.\n"
            f"Actively avoid English sentence patterns; restructure sentences to sound natural and idiomatic in {dest_lang}.\n"
            f"Prefer concise, clear, neutral, and readable phrasing suitable for fast on-screen reading.\n"
            f"Avoid stiffness, redundancy, and word-for-word translation.\n"

            f"Selection and output rules:\n"
            f"Only translate lines that start with 'Line ' followed by a number and a semicolon.\n"
            f"For each such line, translate only the TEXT after the first semicolon.\n"
            f"After translation, do NOT include 'Line N:' in the output; only output the translated TEXT.\n"
            f"Output only {dest_lang} text, with no explanations or comments.\n"
            f"Produce exactly {len(lines)} output lines, in the same order as the input; there should be no blank lines.\n"
            f"End each output line with a single newline character (LF) except for the last line; do not add extra blank lines.\n"

            f"Verb tense handling:\n"
            f"When translating verb tenses, prefer simple, natural, and commonly used {dest_lang} tenses; avoid unnecessary preservation of English tense complexity unless it carries essential meaning.\n"

            f"Handling idioms and cultural references:\n"
            f"Apply adaptation only when an idiom or reference would be unclear or misleading if translated literally.\n"
        )

        # ─────────────────────────────────────────────
        # Persian-specific rules
        # ─────────────────────────────────────────────
        if dest_lang.lower() == "persian":
            prompt = (
                f"SYSTEM IDENTITY & MISSION\n"
                f"[ROLE]: Elite_EN2FA_Translator&Architect + Senior_Persian_Editor (Supreme Master TV Standard)\n"
                f"[TARGET]: Standard Written Persian (Native Broadcast; Subtitle-Ready; Warm/Concise; Genre-Adaptive) (فارسی معیار نوشتاری، روان و ایجازمحور، غیرمحاوره‌ای و قابل‌خواندن در زیرنویس)\n"
                f"[METHOD]: ISO 17100-style;\n"
                f"Agentic reasoning (silent analysis -> reflexion -> output)\n"
                f"[OUTPUT]: Raw_Text_Only | NO Meta/Markdown\n"
                f"\n"
                f"DEFINITIONS (INTERNAL; FOR CONSISTENCY)\n"
                f"\n"
                f"LINE == exactly one input line.\n"
                f"\"segment\" == the same LINE. (No merge/split; ۱:۱)\n"
                f"<L_n> == internal masking placeholder. MUST be restored byte-identical in STEP 5;\n"
                f"MUST NOT appear in final output.\n"
                f"<STANDARD_X> == internal terminology-memory key. MUST NOT appear in final output.\n"
                f"PRIORITY KERNEL (HARD PRECEDENCE)\n"
                f"P0 (IMMUTABLE) > P1 (PRECISION) > P2 (TECHNICAL) > P3 (STYLE).\n"
                f"Violation of Higher_P to satisfy Lower_P == FORBIDDEN.\n"
                f"\n"
                f"F) SUPREMACY CLAUSE (CONFLICT RESOLUTION — BINDING)\n"
                f"[HIERARCHY]: P0 (Safety) >> P1 (Precision) >> P2 (Technical) >> P3 (Style).\n"
                f"[PERSIAN_TRUST]: Existing Persian = trusted UNLESS violates P0 safety or creates logical impossibility.\n"
                f"[TECH_PRIORITY]: P2-A(MustKeepAsIs) >> P1(Digits/Compression).\n"
                f"[SPECIFICITY]: P2-C(Pedagogical) >> P2-G(Latin_Leakage).\n"
                f"[SAFETY_NET]: IF (Rule_Application breaks P0) -> REVERT(Byte-Identical).\n"
                f"[FAILSAFE]: IF Uncertainty_Exists -> FLAG \" ⚠️\".\n"
                f"[RULESET: THE LAWS]\n"
                f"\n"
                f"P0 — SAFETY & LANG (IMMUTABLE)\n"
                f"\n"
                f"LANG: Standard Written Persian ONLY (except allowed P2-A/P2-C passthrough tokens).\n"
                f"P1 — LOGIC & NUMBERS (PRECISION)\n"
                f"\n"
                f"TECH_LOCK: IDs/Versions/Codes (e.g., H5N1, v2.1, ISO-9001) -> IMMUTABLE.\n"
                f"SCOPE_LOCK: Negation/contrast/exception/focus scope -> exact.\n"
                f"[NUMERIC_SEQUENCE] (strict order):\n"
                f"\n"
                f"NUM_CONVERT: Convert written-out numbers to digits (e.g., \"ten\" -> 10).\n"
                f"COMPRESS: Simplify large numbers before localization: \"X,000\" -> \"X هزار\" |\n"
                f"\"X,000,000\" -> \"X میلیون\"\n"
                f"DIGITS: Localize ALL remaining digits (0-9 -> ۰-۹). No rounding. Format: Thousands=\"٬\" ; Decimal=\".\" |\n"
                f"Example: 12,345.67 -> ۱۲٬۳۴۵.۶۷\n"
                f"[UNITS] (VIEWER-FIRST):\n"
                f"IF (Imperial + Metric) -> KEEP Metric ONLY.\n"
                f"ELSE (Imperial Only) -> KEEP Value + Translate Unit Name.\n"
                f"[TERMINOLOGY_CONSISTENCY] (Cross-Session):\n"
                f"Algorithm:\n"
                f"\n"
                f"TRACK: IF English term repeats 2+ times -> LOG first Persian translation as <STANDARD_X>.\n"
                f"ENFORCE: On subsequent occurrences -> USE <STANDARD_X> (byte-identical).\n"
                f"DOMAIN_ALIGN: IF term belongs to domain vocabulary (Environment/Spiritual/Animal/Science) -> apply SMTV-appropriate default UNLESS user pre-inserted Persian overrides.\n"
                f"Exception: Context-dependent terms (e.g., \"bank\"=بانک/ساحل) -> rely on SEMANTIC_REPAIR logic.\n"
                f"[PERSIAN_INPUT_PRIORITY]:\n"
                f"Existing Persian = trusted by default.\n"
                f"IF creates logical impossibility (ontological conflict) -> apply SEMANTIC_REPAIR.\n"
                f"IF violates P0 safety (non-Persian script mixed incorrectly) -> standardize.\n"
                f"Exception: Never \"correct\" stylistic choices (both شخص-الاغ and الاغ-شخص valid unless glossary specifies).\n"
                f"\n"
                f"[SMTV_CONVENTION]:\n"
                f"Animals + human terms = respectful standard.\n"
                f"EN→FA: \"bird-citizen\" -> \"شهروند-پرنده\" | \"donkey-person\" -> \"شخص-الاغ\"\n"
                f"FA→FA: Keep as-is (both orders valid unless glossary specifies).\n"
                f"[SEMANTIC_REPAIR]:\n"
                f"IF pre-inserted Persian creates context mismatch (spiritual→physical, animate→inanimate) -> reverse-engineer English source -> output contextual synonym.\n"
                f"Example: \"اعتکاف یخچال\" (glacier retreat wrongly glossed) -> \"عقب‌نشینی یخچال\"\n"
                f"\n"
                f"[RISK_TRIGGERS]:\n"
                f"IF input contains: Numbers/Digits, Percentages, Currency, Dates, Decimals, Unit conversions\n"
                f"-> Flag line for extra validation in STEP 4.\n"
                f"\n"
                f"P2 — TYPOGRAPHY & LATIN GATE (TECHNICAL)\n"
                f"\n"
                f"P2-A MUSTKEEPASIS (Pattern Recognition):\n"
                f"\n"
                f"URLs, Handles (@user), Hashtags (#tag).\n"
                f"News_Sources: Capitalized text inside parentheses at end-of-segment or end-of-line. Examples: \"(Reuters)\", \"(Al Jazeera)\", \"(Phys.org)\", \"(Thanh Niên)\" [CONSTRAINT]: Keep source name byte-identical (Latin).\n"
                f"P1 Tech-Lock codes/IDs/versions. Action: KEEP AS IS (byte-identical).\n"
                f"P2-C PEDAGOGICAL_PASSTHROUGH:\n"
                f"WHEN: Teaching/definition context + foreign text/word present.\n"
                f"DO: PRESERVE foreign text AS-IS (byte-identical) + translate explanation only.\n"
                f"FORMAT: Foreign_Text یعنی \"Meaning\" OR به Language می‌گویند Foreign_Text\n"
                f"EXAMPLE: \"We say 'Hello' in English\" -> به انگلیسی می‌گویند 'Hello' (یعنی \"سلام\")\n"
                f"\n"
                f"P2-G LATIN_LEAKAGE:\n"
                f"\n"
                f"Output Latin chars [A-Z a-z] == 0\n"
                f"Exceptions: P2-A and P2-C only.\n"
                f"ACRONYM_STRATEGY (No Latin output):\n"
                f"\n"
                f"IF Phonetic (NASA/UNESCO) -> Transliterate (ناسا/یونسکو).\n"
                f"IF Initialism (UN/EU/WHO) -> Translate to standard Persian (سازمان ملل/اتحادیه اروپا/سازمان جهانی بهداشت).\n"
                f"CLEANLINESS:\n"
                f"\n"
                f"Strip harakat (unless needed for disambiguation).\n"
                f"Ezafe: silent \"ه\" -> \"هٔ\" (خانهٔ). FORBID \"ه‌ی\".\n"
                f"Remove Oxford comma before \"و\".\n"
                f"SPACING: Fix ZWNJ where required (می‌شود، کتاب‌ها، خانهٔ).\n"
                f"[SYM_SYNC]:\n"
                f"\n"
                f"NEVER output \"؛\" unless \";\" exists in source; spontaneous injection is FORBIDDEN.\n"
                f"Quotes (SMART_QUOTING):\n"
                f"\n"
                f"STANDARD (Outer): Always use ASCII double quotes (\"...\") as the primary layer for all quotes, dialogue, and emphasis.\n"
                f"NESTED (Inner): If a quote exists inside a standard quote, use Persian guillemets («...»).\n"
                f"SAFETY_MARKER: Use '...' for names lacking established transliteration in mainstream Persian media (e.g., 'مایک دولینسکی' vs. ایلان ماسک).\n"
                f"[IMMUNITY]: NEVER quote:\n"
                f"Job Titles / Honorifics\n"
                f"Globally famous entities\n"
                f"Names immediately following a Title-Anchor (e.g., استاد، دکتر، آقای، خواهر).\n"
                f"P3 — PERSONA, HONORIFICS & STYLE (CONTROLLED)\n"
                f"\n"
                f"MODE RULES:\n"
                f"\n"
                f"SAGE: (1st/2nd Person + Spiritual) warm/reverent; allow Verb_Distance<=12;\n"
                f"Split long sentences if needed in same line.\n"
                f"NEWS: objective/concise; prefer Verb_Distance<=8 when feasible; neutral verbs (گفت/اعلام کرد).\n"
                f"FACILITATOR: warm/standard;\n"
                f"clarity first.\n"
                f"HONORIFIC_PROTOCOL (Non-Inventive, Consistent):\n"
                f"\n"
                f"IF referent explicitly has titles (Master/Supreme Master/Her Holiness/His Holiness/Your Majesty): -> use respectful Persian title + plural honorific verbs (فرمودند/گفتند/تأکید کردند).\n"
                f"IF referent explicitly == God/Allah/Lord: -> reverent lexicon; typically singular verb (می‌فرماید/فرمود).\n"
                f"Do not switch (فرمودند <-> گفت) arbitrarily for same speaker in contiguous block.\n"
                f"DE-TRANSLATIONESE (SAFE):\n"
                f"\n"
                f"Avoid \"توسط\" when meaning preserved.\n"
                f"Max 2 \"که\" per sentence; rewrite if exceeded.\n"
                f"[PRO_DROP_PREFERENCE]: Prefer dropping redundant pronouns (من، ما، او، تو) at sentence start ONLY when unambiguous.\n"
                f"Keep pronouns whenever needed for clarity, emphasis, contrast, or natural Persian rhythm.\n"
                f"[DYNAMIC_VERBS]: Aggressively compress bureaucratic compounds: \"مورد بررسی قرار داد\" -> \"بررسی کرد\" \"به انجام رساند\" -> \"انجام داد\"\n"
                f"RHETORICAL_OVERRIDE (SAGE only, optional):\n"
                f"\n"
                f"If source starts with strong \"Emotional Hook\" or \"Philosophical Axiom\", you are AUTHORIZED to delay the verb (RHYTHM > SYNTAX) to preserve impact.\n"
                f"Constraint: grammatically valid Persian; meaning unchanged; never override P0-P2.\n"
                f"PATTERNS (Common Transformations):\n"
                f"\n"
                f"[A] Introductions:\n"
                f"\"I'm [X]\" / \"My name is [X]\" -> \"من [X] هستم\"\n"
                f"\n"
                f"[B] Greetings:\n"
                f"\"Welcome to [X]\" -> \"خوش آمدید به [X]\"\n"
                f"\n"
                f"EXECUTION (SILENT; PER LINE; ONE-PASS; STRICT)\n"
                f"CONST MAX_RETRY = 1.\n"
                f"\n"
                f"PRE-FLIGHT (SILENT — NO OUTPUT)\n"
                f"\n"
                f"Detect MODE (NEWS/SAGE/FACILITATOR)\n"
                f"Check RISK_TRIGGERS -> Set VALIDATION_LEVEL (STRICT/STANDARD) [NEVER output this analysis]\n"
                f"[STEP 0 / LOCK_MAP & MASK]:\n"
                f"Action A (Technical Locking):\n"
                f"Identify immutable spans: P1_TECH_LOCK + P2-A (URLs/Handles/Hashtags/Sources) + P2-C (foreign pedagogical terms).\n"
                f"-> Mask as <L_n> to protect from modification.\n"
                f"\n"
                f"Action B (Semantic Preservation):\n"
                f"Identify honorific referents per P3 HONORIFIC_PROTOCOL (Master/God/Allah/Lord/titles).\n"
                f"-> Keep visible (do NOT mask) to enable correct verb agreement in Step 3.\n"
                f"\n"
                f"[STEP 1 / SCAN & MODE]:\n"
                f"A. Detect BASE_MODE:\n"
                f"IF News_Markers (dateline/report/official statements/agency/number-heavy) -> MODE=NEWS\n"
                f"ELSE IF Spiritual_Markers (Master/God/Soul/Heaven/Prayer/Meditation + direct address) -> MODE=SAGE\n"
                f"ELSE MODE=FACILITATOR\n"
                f"[FAILSAFE_MODE]: IF unable to detect -> DEFAULT to FACILITATOR (safest/most neutral).\n"
                f"B. Apply ANTI-JITTER:\n"
                f"TRIGGER any of:\n"
                f"\n"
                f"Short_Line: < 4 words\n"
                f"Backchannel: yes/no/ok/right/sure/thanks/I see\n"
                f"Fragment: No explicit subject or main verb (e.g., \"And then?\", \"Which one?\", \"Never.\") ACTION: -> INHERIT MODE + Register from Previous_Non_Ambiguous_Line.\n"
                f"Prevent arbitrary register switches inside one speaker block.\n"
                f"[STEP 2 / DRAFT]:\n"
                f"[PRE-FLIGHT]: Scan session memory for LOCKED_TERMS (<STANDARD_X> mappings from previous lines).\n"
                f"IF input contains Persian text:\n"
                f"[Check A] Safety: Does it violate P0 (non-Persian script mixed incorrectly)? -> Standardize.\n"
                f"[Check B] Semantic: Does it create logical impossibility (per SEMANTIC_REPAIR)? -> Fix.\n"
                f"[Check C] Terminology: Is it a known term?\n"
                f"-> LOCK as <STANDARD_X> for cross-line consistency.\n"
                f"[Check D] Trust: Otherwise -> PRESERVE existing Persian (byte-identical).\n"
                f"THEN: Translate remaining English applying P0+P1+P2 (keep <L_n> intact).\n"
                f"[CONSISTENCY_ENFORCE]: IF English term matches previous occurrence -> REUSE <STANDARD_X> (byte-identical).\n"
                f"[STEP 3 / STYLE]:\n"
                f"Apply P3 (persona/honorifics/pro-drop/rhetorical).\n"
                f"\n"
                f"[STEP 3.5 / BYTE-PRESERVE SHIELD (INTERNAL; OPTIONAL BUT CONSISTENT WITH RULES)]\n"
                f"If any Persian span was PRESERVED (byte-identical) in STEP 2 [Check D], and later steps could alter it (diagnostic rewrite or STEP 5 polish):\n"
                f"-> Temporarily mask that preserved span as <L_n> AFTER STEP 3 (so honorific agreement is already resolved).\n"
                f"(Then STEP 5 Action 2 restores it byte-identical.)\n"
                f"\n"
                f"[STEP 4 / DIAGNOSTIC — HARD]:\n"
                f"Check hard errors: digits/scope/Latin leakage.\n"
                f"[DIGIT_CHECK] (if RISK_TRIGGERS flagged):\n"
                f"ALL Latin digits [0-9] converted to Persian (۰-۹)? (except inside <L_n>)\n"
                f"Numbers ≥1,000 have \"٬\" separator?\n"
                f"Violation -> HARD_ERROR.\n"
                f"\n"
                f"[LATIN_CHECK]:\n"
                f"Latin chars [A-Z a-z] allowed ONLY inside <L_n>; found outside -> HARD_ERROR.\n"
                f"IF ERROR -> REWRITE (decrement MAX_RETRY).\n"
                f"\n"
                f"[STEP 5 / POLISH & RESTORE]:\n"
                f"Action 1: Fix Punctuation, ZWNJ, Ezafe(هٔ), spacing on visible Persian text (ignore <L_n>).\n"
                f"Action 2: Restore all <L_n> placeholders to exact original bytes (byte-identical).\n"
                f"[STEP 6 / FINAL_GATE — BLIND TEST] (MUST PASS):\n"
                f"[BALANCE_CONSTRAINT]: Q1 (Clarity) and Q4 (Precision) must coexist.\n"
                f"Simplification is allowed ONLY if meaning is preserved.\n"
                f"\n"
                f"Q1 (شفافیت / Clarity): \"اگر مخاطب فقط این متن زیرنویس فارسی را بخواند، آیا معنا را در «همان لحظه اول» (بدون نیاز به مکث) کاملاً درک می‌کند؟\"\n"
                f"Q2 (اصالت / Authenticity): \"آیا چیدمان کلمات (Syntax) کاملاً فارسی است یا هنوز ردپای گرامر انگلیسی (بوی ترجمه) در آن حس می‌شود؟\"\n"
                f"Q3 (لحن / Tone): \"آیا واژگان انتخاب‌شده با بافت (خبر/عرفان/گفتگو) و جایگاه گوینده (Sage/News/Facilitator) کاملاً سازگار است؟\"\n"
                f"Q4 (دقت معنایی / Precision): \"آیا تمام جزئیات و ظرایف معنایی متن مبدأ در این ترجمه حفظ شده‌اند یا چیزی فدای ساده‌سازی شده است؟\"\n"
                f"[ACTION]: IF (Q1|Q2|Q3|Q4 == FAIL) AND (MAX_RETRY > 0) -> GOTO STEP 2 (Rewrite).\n"
                f"[FAILSAFE]: IF (Still_Fail) -> FORCE_OUTPUT(Best_Draft + \" ⚠️\").\n"
                f"\n"
                f"REVIEW_LOGIC (APPEND \" ⚠️\" IF ANY)\n"
                f"Inference/forced assumption, unresolved ambiguity, no standard equivalent, risk of violating P0/P1/P2,\n"
                f"uncertain conversion, unclear referent.\n"
                f"[EXIT_PROTOCOL]: STOP -> OUTPUT Final_String only.\n"
                f"\n"
                f"OUTPUT CONSTRAINTS (HARD)\n"
                f"N_input_lines == N_output_lines (Strict Sequencing).\n"
                f"MERGE/SPLIT == FORBIDDEN.\n"
                f"INTEGRITY: Output line count != Input line count = {len(lines)} -> INVALID_RESPONSE.\n"
                f"INTEGRITY_RETRY: Before emit, verify N-line sync;\n"
                f"if mismatch, silently repair and re-check until equal.\n"
                f"\n"
                f"Here is the input text to translate:\n"
            )
        # ─────────────────────────────────────────────
        # Text payload
        # ─────────────────────────────────────────────
        prompt += f"\nHere is the text to translate:\n{numbered_text}\n"

        return prompt

    def translate(self, source_lang, dest_lang, text_to_translate):
        # Auto-create doc ID & filename if missing
        if not self.doc_id:
            print("[INFO] No document context. Generating new doc_id and using filename 'inline_text'.")
            self.set_filename("inline_text")

        prompt = self.build_translation_prompt(source_lang, dest_lang, text_to_translate)
        print("prompt:")
        print(prompt)

        input_tokens = self.estimate_tokens(prompt)
        print(f"Estimated number of input tokens: {input_tokens}")

        start_time = time.time()
        if "pro" in self.model:
            response = self.client.responses.create(
                model=self.model,
                input=[
                    {"role": "system", "content": "You are a professional translator."},
                    {"role": "user", "content": prompt}
                ]
            )
        else:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a professional translator."},
                    {"role": "user", "content": prompt}
                ]
            )
        elapsed_time = time.time() - start_time
        response_json = response.model_dump()

        print("response:")
        print(json.dumps(response_json, indent=4))
        print("--end of response--")

        cost_info = self.calculate_openai_cost(response_json)

        translated_text = response.choices[0].message.content.strip()
        
        # Remove duplicate new lines if any
        translated_text = re.sub(r'\n+', '\n', translated_text)

        # Validate line counts
        in_lines = text_to_translate.split("\n")
        out_lines = translated_text.split("\n")

        if len(in_lines) != len(out_lines):
            print("[WARNING] Line count mismatch!")
            print(f"Input lines: {len(in_lines)}, Output lines: {len(out_lines)}")
            out_lines = "\n".join(out_lines)
            if len(out_lines) > len(in_lines):
                print("Error in openai translation, too many lines")
            else:
                print("Error in openai translation, too few lines")
            translated_text = out_lines

        # Save query record
        try:
            conn = self.get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO queries
                (doc_id, model_name, prompt_json, response_json, execution_time_sec, input_tokens, output_tokens, total_tokens, cost_usd)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    self.doc_id,
                    self.model,
                    json.dumps(prompt),
                    json.dumps(response_json),
                    elapsed_time,
                    cost_info["prompt_tokens"],
                    cost_info["completion_tokens"],
                    cost_info["total_tokens"],
                    cost_info["total_cost_usd"]
                )
            )
            conn.commit()
        except Exception as e:
            print(f"[Warning] Failed to save query info: {e}")
        finally:
            try: cursor.close(); conn.close()
            except: pass

        return response, translated_text
