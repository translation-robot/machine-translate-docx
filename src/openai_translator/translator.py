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
            prompt += (
                f"Language-specific rules for {dest_lang}:\n"
                f"- Numbers: Use a dot for decimals (12.5) and preserve commas for thousands (12,000).\n"
                f"- Thousand rule: Convert “X,000” to “X هزار” ONLY for approximate or descriptive quantities; "
                f"do NOT convert years, IDs, codes, exact prices, or technical measurements.\n"
                f"- Do NOT add diacritics (short vowels or tashkeel), unless a rare word would be ambiguous without them.\n"
                f"- Use phonetic transliteration for foreign terms when there is no well-established {dest_lang} equivalent.\n"
                f"- Use Latin script only for news agency names.\n"

                f"\nMANDATORY NATIVE AUTHOR GATE (FINAL – NON-NEGOTIABLE)\n"
                f"This gate overrides all stylistic intuition and operates on violation detection.\n"

                f"\nFor EACH line, perform the following two-step test:\n"

                f"\nSTEP 1 — COLLOCATION & EXPRESSION VIOLATION SCAN\n"
                f"Ask explicitly:\n"
                f"\"Does this line contain ANY phrase, metaphor, abstraction, or verb–noun pairing "
                f"that would be understandable in Persian "
                f"BUT would NOT naturally appear in modern Persian TV scripts "
                f"written originally in Persian by a professional TV editor?\"\n"

                f"\nRED FLAG VIOLATIONS INCLUDE (but are not limited to):\n"
                f"- Literal or semi-literal English metaphors\n"
                f"- Abstract nouns acting as subjects or objects of vague verbs\n"
                f"- Overuse of nominalizations where Persian prefers verb-based structures\n"
                f"- Expressive but unspecific phrasing that sounds literary or contemplative\n"
                f"- Sentences that feel polished yet lack concrete spoken logic\n"

                f"\nIMPORTANT:\n"
                f"Comprehensibility is NOT a defense.\n"
                f"Expressiveness is NOT a defense.\n"
                f"\"If it sounds fine\" is NOT a defense.\n"

                f"\nIf ANY violation is detected → the line is INVALID.\n"

                f"\nSTEP 2 — NATIVE AUTHOR CONFIRMATION (ONLY IF STEP 1 PASSES)\n"
                f"Ask:\n"
                f"\"Would a professional Persian TV editor, writing directly in Persian with no English source, "
                f"naturally produce this exact structure and phrasing?\"\n"

                f"- If YES with full confidence → keep the line unchanged.\n"
                f"- If NO, uncertain, or requires justification → the line is INVALID.\n"

                f"\nMANDATORY REWRITE RULES (ONLY FOR INVALID LINES)\n"
                f"- Preserve the exact meaning and intent\n"
                f"- Avoid simplification, explanation, or generalization\n"
                f"- Rewrite ONLY by:\n"
                f"  • replacing non-native collocations\n"
                f"  • grounding abstractions into concrete Persian verbs\n"
                f"  • restructuring the sentence into natural Persian SOV flow\n"

                f"DO NOT:\n"
                f"- Keep a sentence because it sounds elegant\n"
                f"- Add interpretive clarity\n"
                f"- Beautify language\n"

                f"\nFINAL QUALITY LOOP (SINGLE PASS)\n"
                f"Before outputting, check each line:\n"
                f"1. Register: Strictly formal broadcast Persian (no broken or spoken forms).\n"
                f"2. Syntax: Fully Persian sentence logic, no English residue.\n"
                f"3. Entity integrity: Names, roles, references, and personhood preserved exactly.\n"
                f"4. Native density: Zero traces of translated cleverness.\n"

                f"If a line passes ALL → output unchanged.\n"
                f"If it fails ANY → rewrite once, then output.\n"
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
