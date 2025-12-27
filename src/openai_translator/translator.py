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
            f"Your task is to translate {source_lang} into {dest_lang}.\n\n"
            f"Overall context and style:\n"
            f"- Read all lines first so you understand the full context.\n"
            f"- Treat the entire input as one coherent text when choosing tone and terminology.\n"
            f"- Ensure consistent translations for recurring terms, names, and concepts across all lines.\n"
            f"- If part of the text to translate is already in {dest_lang}, treat it as a translation memory and keep it literal.\n"
            f"Line-by-line constraints:\n"
            f"- Translate line by line: produce exactly one output line for each input line.\n"
            f"- Do NOT merge, split, add, remove, or repeat lines.\n"
            f"- Use formal, standard and natural {dest_lang} (non-colloquial);preserve all information.\n"
            f"- But the wording inside each line is allowed to become a bit shorter or longer in order to produce natural {dest_lang}.\n"
            f"- Preserve the input line order.\n"
            f"- Parentheses and multiple sentences within a line belong to that same line.\n"
            f"- Only translate lines that start with 'Line ' followed by a number and a semicolon.\n"
            f"- For each such line, translate only the TEXT after the first semicolon.\n"
            f"- After translation, do NOT include 'Line N:' in the output; only output the translated TEXT.\n"
            f"- Output only {dest_lang} text, with no explanations or comments.\n"
            f"- Produce exactly {len(lines)} output lines, in the same order as the input, there should be no blank lines.\n"
            f"- End each output line with a single newline character also known as line feed or LF (do not add extra blank lines between translated lines in the output).\n"
        )
        
        if dest_lang.lower() == 'persian':
            prompt += (
                "- When writing decimal numbers in Persian, use a dot as the decimal separator, e.g. write \u00ab\u06F1\u06F2.\u06F5\u00bb not \u00ab\u06F1\u06F2/\u06F5\u00bb (and not \u00ab\u06F1\u06F2,\u06F5\u00bb).\n"
                "- Do NOT add diacritics (no short vowels or tashkeel such as \u064E \u0650 \u064F \u0651 \u064C \u064B \u064D),  unless a rare word would be ambiguous without them.\n"
                "- Apply the following fixed terminology rules whenever these English forms appear:\n"
                "  - \"animal-person\" / \"animal-people\"  \u2192  \u00ab\u0634\u062E\u0635-\u062D\u06CC\u0648\u0627\u0646\u00bb / \u00ab\u0627\u0634\u062E\u0627\u0635-\u062D\u06CC\u0648\u0627\u0646\u00bb\n"
                "  - \"tiger-person\" / \"tiger-people\"    \u2192  \u00ab\u0634\u062E\u0635-\u0628\u0628\u0631\u00bb   / \u00ab\u0627\u0634\u062E\u0627\u0635-\u0628\u0628\u0631\u00bb\n"
                "  - \"cow-person\" / \"cow-people\"        \u2192  \u00ab\u0634\u062E\u0635-\u06AF\u0627\u0648\u00bb   / \u00ab\u0627\u0634\u062E\u0627\u0635-\u06AF\u0627\u0648\u00bb\n"
                "  (Do NOT translate them as ordinary “animal(s) / tiger(s) / cow(s)”.)\n"
            )
        
        prompt += f"Here is the text to translate:\n{numbered_text}\n"

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
                out_lines = out_lines[:len(in_lines)]
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
