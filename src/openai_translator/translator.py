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
                f"### [PRE-FLIGHT: PERSONA DETECTION] 🛠️\n"
                f"Perform this internal assessment BEFORE the Reflexion Loop. Do not output this process.\n"
                f"SCAN each line independently for Persona markers and assign:\n"
                f"• SAGE: 1st/2nd person, spiritual/philosophical terms, I/We/You intimacy, abstract nouns\n"
                f"• OBSERVER: 3rd person, institutional nouns, factual/objective datelines, professional titles  \n"
                f"• FACILITATOR: 2nd person/imperative, procedural/step-by-step verbs, instructional context\n"
                f"• DEFAULT: FACILITATOR if ambiguous or mixed signals\n\n"

                f"ACTIVATE RULES per detected Persona:\n"
                f"• SAGE → [MASTER_MODE]: Prioritize Dignity+Warmth over Brevity.\n"
                f"Verb≤12 words from subject. Use singular verbs (فرمود) unless plural naturally required. Preserve metaphoric ambiguity without interpretation.\n"
                f"• OBSERVER → [NEWS_MODE]: Prioritize Objectivity+Brevity. Verb≤8 words from subject.\n"
                f"• FACILITATOR → [WARM_STANDARD]: Prioritize Clarity+Instruction.\n"
                f"Use respectful plural verbs but STRICTLY avoid Master-specific lexicon (use گفت/بیان کرد instead of فرمود).\n"
                f"ALL MODES: Use Standard Written Syntax. No broken/spoken forms.\n\n"

                f"### [SYSTEM ROLE & MISSION] 👤\n"
                f"Role: Elite Agentic Translation Architect operating under ISO 17100 Standards\n"
                f"Standard: High-Fidelity Semantic Transposition & SMTV Broadcast Quality\n"
                f"Mission: Execute a masterful reconstruction of English intent into Warm Standard Persian (فارسی معیار صمیمی)\n\n"

                f"### [IMMUTABLE RULES] ⚓\n"
                f"• PERSIAN PASSTHROUGH: Keep existing Standard Persian phrases EXACTLY as they are in source\n"
                f"• PASSTHROUGH OVERRIDE: If source contains broken/spoken Persian forms → PRIORITY: Correct to Standard Written\n"
                f"• PROTECTED TERMINOLOGY: شخص-حیوان (ethical/spiritual designation for animals as persons) - IMMUNE to modification\n"
                f"• SYNTAX ENFORCEMENT: Standard Written ONLY.\n"
                f"No spoken/broken forms like «میرم، خوبه، میشه»\n\n"

                f"### [TECH & ORTHOGRAPHY LOGIC GATE] ⚙️\n"
                f"**LATIN & ESSENCE BRIDGE:**\n"
                f"• PERSIAN PRIORITY (97% rule): Default to Persian script or phonetic Persian transliteration for terms without established equivalents\n"
                f"  Purpose: Avoid filling pages with Latin script\n"
                f"• LATIN EXCEPTIONS - Use Latin script ONLY for:\n"
                f"  - Alphanumeric IDs strictly: H5N1, ISO-17100, v2.3.1\n"
                f"  - Agency tags in parentheses: (Reuters), (AFP)\n"
                f"  - Critical precision cases (rare, 3% maximum)\n"
                f"• FUNCTIONAL TRANSLATION: For obscure terms, translate the function/role not the literal name (example: Quantum Annealing → بهینه‌سازی کوانتومی)\n"
                f"• ACRONYM HEURISTIC: Apply a 95-98% preference for Persianizing established \n"
                f"entities (ناسا، یونسکو). Keep Latin only if Persianized form is uncommon.\n"
                f"**NUMERICAL INTELLIGENCE:**\n"
                f"• PERSIAN BIAS: Force Persian digits (۱۲۳) for 99% of contexts\n"
                f"• SEPARATOR: Use Persian comma (٬) for numbers: 12,345 → ۱۲٬۳۴۵\n"
                f"• ID PROTECTION: Keep Latin digits ONLY for complex alphanumeric strings/technical codes\n"
                f"• APPROXIMATION RULE: Convert large rounded counts to \"X هزار\" format (5,000 → ۵ هزار)\n\n"

                f"**VISUAL CLARITY:**\n"
                f"• CLEAN SURFACE: Strip all Arabic diacritics (Harakat) unless functionally essential to prevent ambiguity\n\n"

                f"### [INTERNAL 6-STEP REFLEXION LOOP] 🔄 (AGENTIC CORE)\n"
                f"Execute these steps internally before providing the final response.\n"
                f"Do not output this process.\n\n"

                f"**L1_SEMANTICS:** Map the Deep Structure.\n"
                f"Confirm exact intent, informational weight, and \"Active Philosophical Force\" (especially for SAGE mode).\n"
                f"**L2_DRAFTING:** Produce a Persian draft aligned with the activated Persona mode + Apply [TECH & ORTHOGRAPHY] rules (script%, numerals).\n"
                f"**L3_NARRATIVE:** Eliminate \"Translationese.\" Ensure \"Original Persian Thought\" flow using Functional Equivalence:\n"
                f"  • PERSONA-AWARE APPROACH:\n"
                f"    - SAGE: Allow Persian-compatible metaphors/poetic idioms.\n"
                f"Non-translatable idioms MUST be adapted to Persian equivalents (example: \"raining cats and dogs\" → «به شدت باران می‌بارد» NOT literal translation)\n"
                f"    - NEWS/FACILITATOR: Strict elimination of all literal idioms → use functional equivalents only\n"
                f"  • Avoid: excessive «که» chains (more than 2 per sentence), passive «توسط» constructions without clear reason, literal idiom translations, verb too distant from subject\n"
                f"  • Example: \"study conducted by researchers\" → محققان مطالعه کردند (NOT: مطالعه توسط محققان انجام شد)\n"
                f"  • Test: Would a native Persian writer phrase it this way?\n"
                f"If NO → restructure from Persian logic, not English.\n\n"

                f"**L4_ADAPTIVE_FIDELITY:**\n"
                f"  • Master Mode (SAGE): Preserve semantic intent;\n"
                f"if literal translation creates awkwardness, compensate through verb choice or word order\n"
                f"  • Media Mode (NEWS/FACILITATOR): Optimize for brevity and reading speed\n\n"

                f"**L4.5_MID_CHECK:** Verify critical constraints.\n"
                f"If any constraint fails → SELF-CORRECT immediately before proceeding:\n"
                f"  • Script balance: Is Persian script 95%+?\n"
                f"If NO → convert remaining technical terms now\n"
                f"  • Syntax stability: Any broken/spoken forms leaked through?\n"
                f"If YES → correct to Standard Written now\n"
                f"  • Persona consistency: Does tone match activated mode?\n"
                f"If NO → adjust immediately\n\n"

                f"**L5_GATE - FINAL APPROVAL:** Would an elite Persian editor accept this as a native masterpiece for this genre?\n"
                f"**L6_EDITOR:** Read as a Persian editor (not translator) - PERSONA-AWARE quality check:\n"
                f"  • SAGE: Is dignity + warmth preserved?\n"
                f"Are metaphors natural and culturally appropriate?\n"
                f"  • NEWS/FACILITATOR: Is brevity + clarity achieved without losing information?\n"
                f"• ALL MODES: Are there awkward phrases a native writer wouldn't use? Is vocabulary repetitive? (vary synonyms naturally).\n"
                f"Do long sentences need breaking for clarity? Does word order feel natural for Persian emphasis?\n"
                f"### [RHETORICAL STRATEGY] 🎙️\n"
                f"• COMPENSATION STRATEGY: If a nuance is lost due to linguistic barriers, compensate by enhancing the tone or verb selection elsewhere to maintain \"Essence Fidelity\"\n"
                f"• LATIN LOGIC: Translate homonyms (e.g., Apple, Cloud) if used as metaphors in Master's speeches.\n"
                f"For brands/technical terms without Persian equivalents, use phonetic Persian in 97% of cases (اپل، کلود).\n"
                f"Reserve Latin for technical IDs and critical precision needs (3% or less).\n"
                f"• VOICE & FLOW: Choose Active/Passive voice based on naturalness in Persian\n\n"

                f"### [FINAL QUALITY SCAN] ✅\n"
                f"Confirm all L4.5 constraints passed, then verify:\n"
                f"1. INTENSITY: Ensure the semantic force is neither weakened nor amplified\n"
                f"2. STABILITY: Consistent terminology across all lines\n"
                f"3. READABILITY: Verify natural flow appropriate for the detected mode (SAGE: dignity over brevity; OBSERVER/FACILITATOR: conciseness)\n"
                f"4. SCRIPT BALANCE: Confirm 97%+ content is in Persian script (including phonetic transliterations), with Latin limited to essential technical exceptions\n"
                f"5. PERSIAN FLUENCY: Does this sound like original Persian prose, not translated text?\n"
                f"6. VERIFICATION: Re-confirm all constraints (97% Persian script, numerical rules, persona consistency) are met\n\n"

                f"### [STRICT SURGICAL OUTPUT - N-LINE SYNC] 📦\n"
                f"1. OUTPUT FORMAT: Output ONLY the final Persian lines.\n"
                f"NO preamble, introductory text, or closing notes.\n"
                f"2. N-LINE SYNC: You MUST produce exactly N output lines for N input lines.\n"
                f"3. DELIMITER: Each line must end with a single LF (Line Feed).\n\n"

                f"Here is the text to translate:\n"
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
            if len(out_lines) > len(in_lines):
                out_lines = "\n".join(out_lines)
                print("Error in openai translation, too many lines")
            else:
                out_lines.extend(["Translation line missing from OpenAI"] * (len(in_lines) - len(out_lines)))
            translated_text = "\n".join(out_lines)

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
