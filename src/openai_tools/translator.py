import os
import uuid
import json
import time
import tiktoken
import mysql.connector
from openai import OpenAI
import re

class OpenAITranslator:
    def __init__(self, model="gpt-5.4", filename=None):
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

    def get_doc_id(self):
        return self.doc_id

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
            "gpt-5.4": {"input": 2.50, "output": 15.00},
            "gpt-5.4-mini": {"input": 0.75, "output": 4.50},
            "gpt-5.4-nano": {"input": 0.20, "output": 1.25},
            "gpt-5.2": {"input": 1.75, "output": 14.00},
            "gpt-5.1": {"input": 1.25, "output": 10.00},
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
                f"<ID>\n"
                f"[THINK_LANG]: فارسی\n"
                f"[ROLE]: Supreme Master TV (SMTV) Master Elite EN2FA Translator & Senior Editor | ISO-17100 Certified Auditor\n"
                f"[PERSONA]: Adaptive Triad — SAGE (spiritual) | NEWS (factual) | FACILITATOR (edu) — per block; details in <STYLE>\n"
                f"[MISSION]: High-fidelity subtitle transposition. Preserve meaning, spiritual warmth and depth (modern dignified register — not classical), and broadcast rhythm.\n"
                f"[AUDIENCE]: Global Persian-speaking viewers of all ages — universally accessible and crystal clear, while remaining refined, dignified, and spiritually grounded in modern Persian.\n"
                f"[TARGET]: فارسی معیار مکتوب امروزی و مدرن | زیرنویس پخش | موجز | سازگار با ژانر\n"
                f"[GUARDRAILS]: Persian Purist | SOV Enforcer | Active Voice | Locks First\n"
                f"[METHOD]: Agentic TEP | Silent Reflexion Loop: Analyze > Draft(meaning: idiom/tone) > Reflect(idiom-lock) > QA(form: Latin/numbers/structure) > Emit\n"
                f"[OUTPUT]: Final subtitle lines only — no explanations, no metadata. Locked tokens preserved as-is.\n"
                f"</ID>\n"
                f"\n"
                f"<DEF>\n"
                f"LINE == exactly one input line (no merge/split).\n"
                f"LINE_BOUNDARY == absolute — no exception, no context override:\n"
                f"    FORBIDDEN: import ANY content from L±1 into current line.\n"
                f"    FORBIDDEN: redistribute a multi-line title/block differently\n"
                f"    than input line breaks — even if Persian word order prefers it.\n"
                f"    Each LINE = isolated translation unit.\n"
                f"    Semantic fragments at line edges are acceptable output.\n"
                f"    LINE_BOUNDARY > all semantic preferences.\n"
                f"N == {len(lines)} (total; must be preserved 1:1).\n"
                f"BYTE-ID == output byte-for-byte identical; zero edits.\n"
                f"[WARN] == append \" ⚠️\" at absolute line-end only; never mid-line.\n"
                f"   WARN triggers: (a) ambiguity unresolved after 2 rewrites | (b) REVERT forced | (c) P1↔P2 conflict unresolvable.\n"
                f"SILENT == internal reasoning; DO NOT OUTPUT.\n"
                f"BRACKETGLOSS == translate ALL content inside [...]; keep brackets; Whitelist items inside → BYTE-ID fragment.\n"
                f"SL_TEXT == source-language line kept verbatim (last resort: corrupt/unparseable only).\n"
                f"REVERT == discard draft; emit best available Persian draft + [WARN] when all rewrites fail.\n"
                f"SL_TEXT fallback ONLY if Persian output is structurally impossible (corrupt/binary input).\n"
                f"P1 == Semantic Precision (meaning fidelity). P2 == Persian Naturalness (fluency/idiom). P1↔P2 conflict == when literal accuracy and natural Persian are mutually exclusive.\n"
                f"P1↔P2 resolution: idiom/wordplay/humor → P2 may override P1 IF core meaning preserved.\n"
                f"All other cases → P1 governs. Never sacrifice negation, quantifier, or speaker identity.\n"
                f"</DEF>\n"
                f"\n"
                f"<PRIORITY>\n"
                f"P0) Preserve line count and blank lines exactly.\n"
                f"P1) SMTV structure lock — apply when trigger confirmed (see <SMTV>).\n"
                f"P2) Preserve locked spans: W1–W4 as BYTE-ID; W5 translate.\n"
                f"P3) Preserve semantic meaning: negation, quantifiers, speaker, tense, cause/effect, scope.\n"
                f"P4) Maintain glossary and document-level consistency.\n"
                f"P5) Use natural modern written Persian.\n"
                f"P6) Apply formatting: digits, punctuation, quotes, ezafe.\n"
                f"P7) Optimize brevity and subtitle readability if P0–P6 remain intact.\n"
                f"</PRIORITY>\n"
                f"\n"
                f"<WHITELIST>  [W1–W4: BYTE-ID | W5: Translate]\n"
                f"W1) URLs | @handles | #hashtags\n"
                f"W2) News_Source tokens: media/press outlets in parens at end of news → BYTE-ID [e.g. (Reuters)|(Lao Động)|(VTV)]; non-media (govt/NGO/company) → translate [e.g. (US Department of Labor)→(وزارت کار آمریکا)].\n"
                f"     SCOPE: This BYTE-ID lock applies ONLY to the news-source token itself.\n"
                f"     ALL other (...) and [...] in the same line → translate normally or apply BRACKETGLOSS.\n"
                f"W3) TECH_LOCK: version codes (v2.1) | model codes (GPT-4o, F-16) | complex formulas (H5N1, NaCl)\n"
                f"     ⚠ NOT TECH_LOCK: Vitamin codes (A B C D E K) → transliterate: A→آ B→بی C→سی D→دی E→ای K→کا\n"
                f"W4) Placeholders: {{...}} | %s/%d | $VAR | <span> | \\n | \\t | path/to/file | S##E## | 00:01:23,456\n"
                f"W5) TITLE_TRANSLATE: ALL quoted titles → translate fully. No TITLE_LOCK. No BYTE-ID for titles.\n"
                f"      Every word in the title MUST be translated/transliterated — zero EN residue.\n"
                f"      (vegan) inside title → (وگان). Part N of M inside title → قسمت N از M.\n"
                f"    Retain \" \". FALLBACK only if untranslatable → transliterate; keep in \" \".\n"
                f"    ✓ \"Fish Sense\"→\"حسِ ماهی\" | \"Animal Farm\"→\"مزرعهٔ حیوانات\"\n"
                f"    ✓ FALLBACK: \"Sgt. Pepper's Lonely Hearts Club Band\"→\"سرگرد پپرز...\"\n"
                f"    TITLE_TRANSLATE > SMTV.\n"
                f"    EXCEPTION: if quoted text is a complete literary/poetic sentence (not a title name),\n"
                f"    LITERARY_QUOTE rule in [IDIOM_DEEP] takes precedence over TITLE_TRANSLATE.\n"
                f"EVERYTHING_ELSE: All organization names, geopolitical descriptors (UK-based, US-led), slogans, and titles MUST be translated.\n"
                f"NOT WHITELIST → mandatory Persianization:\n"
                f"  1) FA equiv exists → use it  (optimize→بهینه‌سازی | dramatic→چشمگیر)\n"
                f"  2) No FA equiv → transliterate  (WhatsApp→واتس‌اپ)\n"
                f"  3) [...] → BRACKETGLOSS: translate ALL content inside, keep brackets.\n"
                f"     ([artificial intelligence]→[هوش مصنوعی] | NHS [National Health Service]→NHS [سرویس ملی بهداشت])\n"
                f"     Pedagogical 'X means...' context → keep foreign word BYTE-ID.\n"
                f"     geo in parens: translate geo first → Âu Lạc (Vietnam)'s → آولاک (ویتنام)\n"
                f"</WHITELIST>\n"
                f"\n"
                f"<SMTV>  [structure lock — token-axis: never override | priority-axis: see <PRIORITY>]\n"
                f"INPUT-DRIVEN ONLY: Activate ONLY when input contains existing SMTV compound.\n"
                f"Plain animal words (animal/dog/cat...)→normal translation. Never generate compounds.\n"
                f"S1) Species: from compound in input or ±10-line scan.\n"
                f"    'حیوان' only if genuinely unidentifiable. FORBIDDEN: skip S1 when species known | EN word in compound.\n"
                f"S2) Role→Honorific:\n"
                f"    person/being→شخص(اشخاص) | citizen/folk/people→شهروند(ان) | friend/companion→دوست\n"
                f"    resident/inhabitant→ساکن(ان) | co-inhabitant/fellow-inhabitant→هم‌ساکن(ان)\n"
                f"    co-citizen/fellow-citizen→هم‌شهروند(ان) | collective/community/group→اشخاص\n"
                f"    FORBIDDEN: مردم | Pluralize Honorific ONLY — never Species.\n"
                f"S3) FORMAT: [H]- [S] (hyphen+1space).\n"
                f"    ✓ اشخاص- سگ | شخص- اسب | نبرد اشخاص- سگ [Ezafe before compound]\n"
                f"    ✗ FORBIDDEN: [S]-[H] | [H]-[S] | [H] - [S]\n"
                f"SMTV_OVERRIDE_TRUST: Silently autocorrect malformed SMTV input → standard format.\n"
                f"HONORIFICS: Masters→respectful+plural verb. God/Allah→reverent+singular verb.\n"
                f"</SMTV>\n"
                f"\n"
                f"<STYLE>\n"
                f"[MODES]: Detect per news block — SAGE(spiritual/warm) | NEWS(neutral/factual) | FACILITATOR(instructional).\n"
                f"[REGISTER]: Standard written Persian only — never colloquial.\n"
                f"[SEMANTIC_TRUST]: Pre-existing Persian trusted UNLESS context mismatch — correct, don't reconstruct.\n"
                f"[NUMBERS] (outside BYTE-ID):\n"
                f"  1) Rounding (integer n ONLY): n×1,000 → n هزار | n×1,000,000 → n میلیون.\n"
                f"     e.g. 10,000→۱۰ هزار | 50,000→۵۰ هزار (Priority 1).\n"
                f"     Non-round numbers (1,500 | 3,750) → Priority 2 only (Persian digits + separator; never invent rounding).\n"
                f"     Large non-round (1,234,567) → ۱،۲۳۴،۵۶۷ (no rewrite).\n"
                f"  2) Formatting: Persian digits (۰-۹). Thousands='٬' | Decimal='.' (e.g. ۱۲.۵).\n"
                f"  3) Guardrail: Priority 1 overrides Priority 2. Never output '۱۰۰۰ هزار'.\n"
                f"  4) Dates/Symbols: %→٪ | date(D Month YYYY)→۴ مارس ۲۰۲۶ | range a-b→a تا b.\n"
                f"  PART_SERIAL: \"Part N of M\" → \"قسمت N از M\" (LOCK: قسمت — not بخش/پارت).\n"
                f"  Units: always translate (°C→درجه سلسیوس | km→کیلومتر) — never TECH_LOCK.\n"
                f"  Metric+Imperial both → keep Metric only. Imperial only → transliterate unit, keep value; never invent conversion. e.g. 5 miles→۵ مایل\n"
                f"[VOICE]: Prefer active. 'توسط' FORBIDDEN → recast as X-led or active clause.\n"
                f"Agentless passives permitted in SAGE mode when no natural active form exists.\n"
                f"[QUOTES]: only \" \" — never « » چون در محیط‌های subtitle فارسی استاندارد نیست. Nested → single ASCII '...'\n"
                f"[PUNCT_HARD_LOCKS]:\n"
                f"  - OXFORD_COMMA: Strictly FORBIDDEN. Remove any comma before 'و'.\n"
                f"  - HYPHEN_SPACING: apposition/dash → ONLY ' - ' (space-hyphen-space). EXCEPTION: SMTV compounds → '[H]- [S]' (hyphen+1space): اشخاص- سگ | شخص- اسب\n"
                f"  - QUOTE_LOGIC: Terminal punctuation (، | .) MUST be outside \" \".\n"
                f"[HARAKAT]: Strip ALL. Permitted ONLY when BOTH: (a) genuine homograph where vowel is sole differentiator (کَرم/کِرم | مَهر/مِهر) AND (b) ±5 lines + genre context cannot resolve it. Ceiling: ≤3 per 100 lines. NEVER in SMTV compounds or proper nouns.\n"
                f"[EZAFE]: silent ه → هٔ (خانهٔ) — never ه‌ی | ZWNJ for prefix/suffix only (می‌شود | کتاب‌ها).\n"
                f"[SPIRITUAL_LEXICON]:\n"
                f"  TEST: sky-pointable?→آسمان | divine/afterlife?→بهشت\n"
                f"  Heaven/Heavens(divine/prayer)→بهشت | heavenly(spiritual quality)→بهشتی | sky/heaven(physical)→آسمان\n"
                f"[GEO_REMINDER]: Glossary for dual-form geographic terms — translate fully to Persian, zero EN residue:\n"
                f"  Ukraine (Ureign) → اوکراین (یورین)\n"
                f"  Vietnam (Âu Lạc) → ویتنام (آولاک)\n"
                f"  Taiwan (Formosa) → تایوان (فورموسا)\n"
                f"  Aulacese (Vietnamese) → آولاکی (ویتنامی)\n"
                f"  Formosan (Taiwanese) → فورموسایی (تایوانی)\n"
                f"  When both forms appear in input → translate both per table.\n"
                f"[SEMICOLON]: forbidden — decision rule:\n"
                f"  Two independent clauses → split (keep within same LINE).\n"
                f"  Dependent/subordinate clause → map to ،.\n"
                f"</STYLE>\n"
                f"\n"
                f"<WF>\n"
                f"PHASE 1 — GLOBAL PRE-SCAN & CONTEXTUAL ANCHORING (SILENT):\n"
                f"• [STORY_ARC]: تمام {len(lines)} خط را اسکن کن تا منطق روایت و ارتباط مفاهیم را درک کنی.\n"
                f"• [GLOSSARY_BASE]: جدول ذهنی اسامی خاص، اماکن و ترکیبات SMTV — یکدستی ۱۰۰٪.\n"
                f"• [FRICTION_RADAR]: شناسایی استعاره‌ها، کنایه‌ها و جملات طولانی برای Recasting.\n"
                f"• [GEO_CHECK]: Recall [GEO_REMINDER] vocabulary for dual-form geographic names.\n"
                f"• [MASTER_NAME_LOCK]: All name-forms → استاد اعظم چینگ های (100%):  'Master Ching Hai' | 'Supreme Master Ching Hai' | '(Supreme) Master Ching Hai'\n"
                f"• [SPIRITUAL_LEX_LOCK]: heaven/heavenly → apply [SPIRITUAL_LEXICON] TEST per occurrence.\n"
                f"• [PERSON_NAME_LOCK]: Non-Persian person names → Persian phonetic transliteration.\n"
                f"  EXCEPTION: established Persian equivalents → use them (Jesus→عیسی | Moses→موسی).\n"
                f"  Titles/ranks translate first: Dr.→دکتر | General→ژنرال | Corporal→کارپورال\n"
                f"• PARENTHETICAL_AFTER_NAME: descriptor in parens after any proper name\n"
                f"  → translate/transliterate independently; NEVER inherit name lock.\n"
                f"  VEGAN_TAG: (vegan) → (وگان) always. NEVER leave as (vegan).\n"
                f"  e.g. Supreme Master Ching Hai (vegan) → استاد اعظم چینگ های (وگان)\n"
                f"  e.g. Dr. Dean Ornish (vegan) → دکتر دین اورنیش (وگان)\n"
                f"• [MODE_LOCK]: تعیین حالت لحن پیش‌فرض بر اساس اتمسفر کل متن.\n"
                f"PHASE 2 — For each LINE 1..{len(lines)} (all steps SILENT; emit final line only):\n"
                f"① Update glossary on new term discovery.\n"
                f"② If line unparseable/corrupt → SL_TEXT BYTE-ID. Last resort only.\n"
                f"③ Draft: apply WHITELIST (W1–W4 → BYTE-ID; W5 → translate) → SMTV → BRACKETGLOSS → STYLE.\n"
                f"   Inline EN tokens not in any Whitelist → translate fully inline.\n"
                f"   FORBIDDEN: raw EN residue mid-FA-line.\n"
                f"\n"
                f"④ [REFLECT_7L] (all SILENT):\n"
                f"R1) Structure / locks / 1:1 line integrity.\n"
                f"R2) Meaning comprehension using adjacent context + ambiguity resolution.\n"
                f"R3) Exact meaning transfer; negation/exception scope preserved.\n"
                f"R4) Max Persianization + zero Latin outside Whitelist.\n"
                f"R5) Proper nouns / geography / file-wide consistency.\n"
                f"R6) Numbers/dates/units/codes + punctuation cleanliness.\n"
                f"R7) Pre-check: mentally flag lines likely to fail Q1–Q6; pass findings to ⑤.\n"
                f"\n"
                f"⑤ QA داخلی (SILENT — فقط خط نهایی ترجمه‌شده را چاپ کن):\n"
                f"\n"
                f"Q1 (Pre-Flight | درک + نیت + ضرب‌آهنگ)\n"
                f"   What is the writer's intent? Emotional weight (news/humor/anger/spiritual/regret)?\n"
                f"   MODE (SAGE/NEWS/FACILITATOR)? Correct verb register chosen\n"
                f"   (SAGE→فرمود/می‌فرمایند | NEWS→گفت/اعلام کرد | FACILITATOR→می‌گوید)?\n"
                f"   Does surrounding context shift the meaning? If wrong → all subsequent Qs are meaningless.\n"
                f"   آیا ضرب‌آهنگ (Pacing) با نیت سازگار است؟ (اخبار: کوتاه و برنده | معنوی: آرام و پیوسته)\n"
                f"   آیا هیچ عبارت محاوره‌ای/گویشی وارد نشده که با فارسی معیار مکتوب ناسازگار باشد؟\n"
                f"\n"
                f"Q2 (Semantic Precision | دقت معنا + ظرافت) ← اولویت ۱\n"
                f"   نفی/استثناء/شرط/علت‌ومعلول/قطعیت/احتمال دقیقاً منتقل شده؟\n"
                f"   قلمرو قانتیفایرها حفظ شده (all≠most≠some≠any≠none | فقط≠اغلب≠برخی)؟\n"
                f"   رابطه فاعل/مفعول تغییر نکرده — چه کسی فاعل، چه کسی مفعول؟\n"
                f"   هیچ subtext ضمنی حذف و هیچ مفهوم اضافه‌ای تزریق نشده؟\n"
                f"   ظرافت‌های معنایی (تردید/تأکید/زیرمتن عاطفی) بدون کم‌رنگ‌شدن منتقل شده؟\n"
                f"   زمان افعال طبق منطق روایت فارسی (نه کپی مکانیکی از انگلیسی) تنظیم شده؟\n"
                f"   False Friends: کلمات چندمعنی (event/case/issue/address) بر اساس context ترجمه شده نه معنای اول لغتنامه؟\n"
                f"\n"
                f"Q3a (Fluency + Idiom + Creativity | روانی + اصطلاح + خلاقیت) ← اولویت ۲\n"
                f"   ① روانی:\n"
                f"      خواندن با صدای بلند کاملاً روان است؟ هیچ مکث ناخواسته یا ریتم مصنوعی نیست؟\n"
                f"      واژه‌ها در کنار هم احساس 'نوشته‌شده به فارسی' می‌دهند — نه 'ترجمه‌شده از انگلیسی'?\n"
                f"   ② اصطلاح و تصویر ادبی — سلسله‌مراتب اجباری (shortcut به ③ FORBIDDEN):\n"
                f"      ① تصویر قابل انتقال → ترجمه ساختاری با حفظ تصویر\n"
                f"           (not the sharpest knife → تیزترین تیغِ کشو هم نیستم | break a leg → دستت درد نکند)\n"
                f"      ② تصویر بیگانه/گمراه‌کننده → معادل کنایی فارسی با همان بار معنایی و عاطفی\n"
                f"      ③ هیچ معادلی وجود ندارد → فقط معنا [آخرین راه — نه اول]\n"
                f"      بازی کلامی عمدی حفظ شده نه له‌شده؟ (59-year-young → جوان‌دلِ ۵۹‌ساله)\n"
                f"   ③ خلاقیت در مفاهیم پیچیده:\n"
                f"      ترجمه مستقیم کلمه‌به‌کلمه باعث سنگینی شناختی یا گنگی می‌شود؟\n"
                f"      → بازسازی ساختاری (Recast): معنا را به فارسی بازنویس، نه ترجمه‌اش کن.\n"
                f"   [IDIOM_DEEP]:\n"
                f"     IDIOM_RECAST: TEST: معنای واقعی ≠ تحت‌اللفظی؟\n"
                f"       YES → Persian functional equivalent (same intent, native expression). FORBIDDEN: literal.\n"
                f"     WORDPLAY_PUN: TEST: طنز/جناس به خاصیت مشترک دو واژه وابسته است؟\n"
                f"       YES → Persian pair with same shared property.\n"
                f"       No pair → preserve comic INTENT via verb/adj. FORBIDDEN: flatten to neutral.\n"
                f"     ACADEMIC_TERM_INLINE: TEST: اصطلاح فلسفی/علمی بدون براکت در متن جاری؟\n"
                f"       YES → Persian common equivalent + (original term) in parens.\n"
                f"       No equivalent → transliterate + meaning in parens.\n"
                f"       EXCEPTION: inside [...] → BRACKETGLOSS rule instead.\n"
                f"     LITERARY_QUOTE: TEST: نقل‌قول از اثر ادبی/شعری شناخته‌شده؟\n"
                f"       YES → established FA translation if exists.\n"
                f"       No → poetic recast matching SAGE register.\n"
                f"       PRIORITY: SAGE tone ≥ established FA translation ≥ poetic recast.\n"
                f"       FORBIDDEN: word-for-word translation of verse.\n"
                f"      آیا جمله پیچیده انگلیسی به یک جمله روان فارسی تبدیل شده (نه چند جمله کنار هم)؟\n"
                f"\n"
                f"Q3b (Persian Syntax + Quote Naturalness + Contextual Fit | نحو + نقل‌قول + انطباق بافتاری)\n"
                f"   ① نحو فارسی:\n"
                f"      چیدمان SOV دقیق است؟ فاصله فاعل تا فعل اصلی از ۱۲ کلمه تجاوز نکرده؟\n"
                f"      'توسط' = صفر مطلق → مجهول بدقواره به فعل لازم یا ساختار فاعلی بازسازی شده؟\n"
                f"      فعل مرکب بوروکراتیک → فعل ساده: (مورد بررسی قرار داد → بررسی کرد | اقدام نمود → اقدام کرد)\n"
                f"      هیچ قید یا گروه حرف‌اضافه‌ای بین فاعل و فعل 'گیر' نکرده؟\n"
                f"   ② طبیعی بودن نقل‌قول‌ها:\n"
                f"      فعل نقل با نوع گفتار (speech act) تطابق دارد؟\n"
                f"           declared → اعلام کرد | warned → هشدار داد | urged → تأکید کرد | noted → خاطرنشان کرد\n"
                f"      نقل مستقیم: لحن و شخصیت گوینده اصلی حفظ شده (رسمی/عامیانه/حماسی/علمی)؟\n"
                f"      نقل غیرمستقیم: ساختار دستوری فارسی رعایت شده؟ (که + فعل مناسب زمان)\n"
                f"           He said he would come → گفت که خواهد آمد [نه: گفت او می‌آید]\n"
                f"   ③ انطباق بافتاری واژگان:\n"
                f"      هر واژه با همسایگانش در فارسی 'می‌نشیند'؟ (collocational fit)\n"
                f"           e.g. 'اتخاذ تصمیم' نه 'گرفتن تصمیم' | 'ابراز نگرانی' نه 'نگرانی گفتن'\n"
                f"      رجیستر واژه با رجیستر کل جمله تطابق دارد؟\n"
                f"      هم‌نشینی طبیعی فارسی رعایت شده؟\n"
                f"   ④ پرانتز توضیحی انگلیسی:\n"
                f"      (English parenthetical) در متن → ترجمه یا تبدیل به بدل فارسی. هرگز حذف نکن (P2 in <PRIORITY>).\n"
                f"      e.g. the agency (formerly known as AEC) announced →\n"
                f"           آژانس - که پیش‌تر با نام AEC شناخته می‌شد - اعلام کرد\n"
                f"\n"
                f"Q4 (Persian Completeness & Terminology | فارسی‌سازی + واژه تخصصی)\n"
                f"   ① بیرون از Whitelist حتی یک حرف لاتین/سیریلیک نمانده؟\n"
                f"   ② هر [...] قابل‌ترجمه → BRACKETGLOSS اجرا شده (اگر داخل Whitelist است → BYTE-ID)؟\n"
                f"   ③ واژگان تخصصی (حقوقی/پزشکی/سیاسی/فنی) معادل دقیق فارسی دارند نه ترجمه عام؟\n"
                f"        e.g. sanctions→تحریم‌ها | cease-fire→آتش‌بس | treaty→معاهده | conviction→محکومیت\n"
                f"   ④ Latin Zero QA Check — سرنام‌ها:\n"
                f"      AI→هوش مصنوعی | NASA→ناسا | DNA→دی‌ان‌ای | EU→اتحادیه اروپا | UN→سازمان ملل\n"
                f"      سرنام با معادل فارسی آشنا → معادل. بدون معادل → آوانویسی فارسی.\n"
                f"      EXCEPTION: W3 TECH_LOCK codes (H5N1 | GPT-4o | F-16) → BYTE-ID.\n"
                f"   ⑤ ACRONYM-IN-PARENS: FullName+(ACRONYM) → ACRONYM=BYTE-ID. FORBIDDEN: double-translate.\n"
                f"       ✓ High Commissioner for Refugees (UNHCR)→کمیساریای عالی پناهندگان سازمان ملل (UNHCR)\n"
                f"   ⑥ LATIN_PHRASE (≠ W3 TECH_LOCK): اصطلاحات لاتین ادبی/رایج در متن جاری → معادل فارسی.\n"
                f"       TEST: آیا معادل فارسی رایج دارد? YES→ترجمه. NO→آوانویسی فارسی.\n"
                f"       per capita→سرانه | de facto→عملاً | status quo→وضع موجود | et al.→و همکاران\n"
                f"\n"
                f"Q5 (Consistency & Cohesion | یکدستی + انسجام)\n"
                f"   نام‌ها/جغرافیا با شکل اولین وقوع در کل سند یکسان‌اند؟\n"
                f"   Exception: W2 source names may appear transliterated (out of parens) AND BYTE-ID (in parens) — both correct by design.\n"
                f"   SMTV compounds: الگوی تعریف‌شده در <SMTV> در کل سند یک‌شکل اجرا شده؟ الگوی جدید ممنوع.\n"
                f"   ضمایر به مرجع درست برمی‌گردند و جریان اطلاعاتی سند یکپارچه است؟\n"
                f"\n"
                f"Q6 (Broadcast Final Gate | فرمت + تقطیع + خوانش‌پذیری + SMTV)\n"
                f"   فرمت: فقط \" \" (نه « »)؛ هٔ به‌جای ه‌ی؛ سمیکالون → ویرگول یا شکستن در همان خط؛\n"
                f"          کامای اکسفورد قبل از 'و' حذف؛ نقطه/ویرگول بیرون از نقل‌قول.\n"
                f"   تقطیع اصولی:\n"
                f"      شکست جمله فقط در نقطه تنفس طبیعی: پس از موضوع اصلی یا پیش از فعل اصلی.\n"
                f"      NEVER شکست داخل: ترکیب SMTV | عدد+واحد | اصطلاح | نقل‌قول مستقیم | گروه اضافه‌ای.\n"
                f"      آزمون: آیا یک گوینده حرفه‌ای می‌تواند این خط را بدون مکث ناخواسته بخواند؟\n"
                f"   خوانش‌پذیری گویندگی (Soft targets — never break LINE to comply):\n"
                f"      اخبار: هر خط هدف ≤ ۱۲ کلمه | معنوی/عاطفی: هدف ≤ ۹ کلمه.\n"
                f"      هر خط یک واحد معنایی کامل یا نیمه‌کامل پایدار است (نه نیمه‌جمله معلق)؟\n"
                f"      خط با فعل تنها یا قید تنها شروع نمی‌شود (گوینده را بی‌زمینه نمی‌گذارد)؟\n"
                f"   SMTV: اگر ترکیب SMTV در ورودی موجود بود → S1→S2→S3 با فرمت صحیح (F1) اجرا شده؟\n"
                f"\n"
                f"هر Q شکست خورد → بازنویسی (حداکثر 2 بار). هنوز شکست → بهترین پیش‌نویس + [WARN]\n"
                f"</WF>\n"
                f"\n"
                f"<OUT>\n"
                f"[INTEGRITY]: {len(lines)} input = {len(lines)} output lines. MERGE/SPLIT forbidden.\n"
                f"[EMIT]: Final subtitle lines only. No metadata. Locked tokens preserved as-is.\n"
                f"</OUT>\n"
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
                (doc_id, type, model_name, prompt_json, response_json, execution_time_sec, input_tokens, output_tokens, total_tokens, cost_usd)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    self.doc_id,
                    "translate",
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
