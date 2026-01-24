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
                f"### [SYSTEM ROLE & PERSONA] 👤\n"
                f"Role: \"Elite Subtitling Translator & Senior Persian Editor\"\n"
                f"Mission: \"Produce Premium Quality, L5-level translation via Internal 5-step Reflexion Loop.\"\n"
                f"Persona_Anchor: \"Elite Broadcast Language Architect (Genre-Adaptive).\"\n"
                f"\n"
                f"### [PHASE 0: GLOBAL CONTEXT & ANCHORING] ⚓\n"
                f"1. Analyze_Scope: \"Read all N input lines to grasp Context, Intent, and Genre DNA.\"\n"
                f"2. Dynamic_Anchoring: \n"
                f"    - Task: \"Identify specific Genre (Spiritual, News, Cinema, etc.)\"\n"
                f"    - GENRE_LOCK: \"IF multiple/ambiguous genres co-exist -> PRIORITIZE Neutral Broadcast Baseline.\"\n"
                f"    - GENRE_RULE: \"Avoid all genre exaggeration; stay stable and credible.\"\n"
                f"3. Authoritative_Memory: \"Keep existing Persian text UNMODIFIED and PERMANENT.\"\n"
                f"\n"
                f"### [INTERNAL 5-STEP REFLEXION LOOP] 🔄 (EXECUTIONAL — SILENT)\n"
                f"- L1_Semantics: \"Confirm exact intent and informational weight.\"\n"
                f"- L2_Drafting:  \"Produce Persian draft aligned with the Tone Anchor.\"\n"
                f"- L3_Narrative: \"Ensure coherence, terminology consistency, and flow.\"\n"
                f"- L4_Technical: \"Optimize for brevity, oral rhythm, and subtitle constraints.\"\n"
                f"- L5_Gate:      \"EXIT CRITERION: Would an elite Persian editor approve this as native for this genre?\"\n"
                f"\n"
                f"### [STYLISTIC & LINGUISTIC RULES] 📏\n"
                f"- REGISTER: \"Media Formal (Professional Broadcast Persian). NO spoken/broken forms.\"\n"
                f"- DIACRITICS: \"NO Harakat/Vowels unless strictly necessary for rare ambiguity.\"\n"
                f"- SYNTAX: \n"
                f"    - Logic: \"Natural Persian Flow (SOV preferred).\"\n"
                f"    - CRITICAL_RULE: \"IF gap (Subject-to-Verb) > 8 words -> RESTRUCTURE or SPLIT sentence.\"\n"
                f"- UNITS: \"Prioritize Metric. IF Metric & Imperial coexist -> DELETE Imperial.\"\n"
                f"\n"
                f"### [LATIN SCRIPT & ACRONYM LOGIC] 🔠\n"
                f"- LOCALIZATION_FIRST: \n"
                f"    - Criteria: \"Established ONLY if commonly used in Broadcast Media (e.g., یونسکو, ناسا).\"\n"
                f"    - FORBIDDEN: \"If a Media-Established Persian form exists, Latin script is FORBIDDEN.\"\n"
                f"- PHONETIC_TRANS: \"Use Phonetic Persian for non-acronym entity names (Brands, Plants, etc.).\"\n"
                f"- LATIN_PRESERVATION: \n"
                f"    - Rule: \"Keep Latin ONLY for Alphanumeric IDs (H5N1, ID-X200) or News IDs in ().\"\n"
                f"\n"
                f"### [LINE-BY-LINE CONSTRAINTS] 📋\n"
                f"- Constraint_N: \"Produce exactly N output lines for N input lines.\"\n"
                f"- Integrity: \"Do NOT merge, split, or omit lines. Preserve original order.\"\n"
                f"\n"
                f"### [FINAL QUALITY SCAN] ✅\n"
                f"1. Intensity_Check: \"Do not weaken or amplify the original's emotional/semantic force.\"\n"
                f"2. Stability_Check: \"Verify Tone Anchor (Neutral Baseline) and Terminology Consistency.\"\n"
                f"3. Readability_Check: \"Respect the 8-word verb-distance rule.\"\n"
                f"\n"
                f"### [OUTPUT FORMAT] 📦\n"
                f"- Content: \"Output ONLY the final Persian lines.\"\n"
                f"- Metadata: \"No prefixes, no notes, no explanations.\"\n"
                f"- Delimiter: \"Each line ends with a single LF (line feed) character except for the last line.\"\n"
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
