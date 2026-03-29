import os
import uuid
import json
import time
import tiktoken
import mysql.connector
from openai import OpenAI
import re

class OpenAISubtitleSplitter:
    def __init__(self, model="gpt-5.4-mini", filename=None, doc_id=None):
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

        if doc_id:
            self.doc_id = doc_id
        else:
            self.doc_id = str(uuid.uuid4())
            
        self.filename = None
        if filename:
            self.filename = filename

    def set_model(self, model):
        """Change openai model for this object."""
        self.model = model
        
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
    def build_subtitle_splitter_prompt(source_lang, dest_lang, source_text, translation):
        lines = source_text.split("\n")
        numbered_lines = [f"Input line {i+1}: {line}" for i, line in enumerate(lines)]
        numbered_text = "\n".join(numbered_lines)
   
        prompt = (
            f"You are an elite video subtitle editor. You are given a subtitle text in {source_lang} (source language) and its translation in {dest_lang} (target language).\n\n"
            
            f"Task:\n"
            
            f"1. Your top priority is that each line in the {dest_lang} translation matches the {source_lang} source line-by-line as closely as possible when viewed side by side.\n"
            f"2. Each line in the {source_lang} source must correspond to exactly one line in the {dest_lang} output.\n"
            f"3. Do not change any words, punctuation, or symbols in the {dest_lang} translation.\n"
            f"4. Preserve all elements such as emails, URLs, symbols, or emojis exactly as they appear in the original {dest_lang} translation.\n\n"
            
            f"5. Only if strict alignment is impossible due to grammar or phrasing differences, apply professional subtitle line-splitting rules for {dest_lang}:\n"
            f"   - Keep grammatical units together (pronouns + verbs, subject + verb, auxiliary + main verb, verb + object, article + noun, preposition + object).\n"
            f"   - Prefer breaking at natural pauses, commas, conjunctions, or between clauses.\n"
            f"   - Do not split fixed expressions, names, phrasal verbs, or idioms.\n"
            f"   - Avoid leaving short words alone (e.g., 'a', 'the', 'to', 'of').\n"
            f"   - Maintain readability and natural flow.\n"
            f"   - Keep line lengths roughly balanced but prioritize meaning and alignment with the source.\n\n"
            
            f"6. Output requirements:\n"
            f"   - Output only the {dest_lang} text.\n"
            f"   - Use exactly {len(lines)} lines, matching the {source_lang} source line count.\n"
            f"   - Insert line breaks only; no paraphrasing, numbering, labels, or extra formatting.\n\n"
            
            f"CRITICAL:\n"
            f"- You MUST output exactly {len(lines)} lines.\n"
            f"- If you cannot split naturally, duplicate or break arbitrarily.\n"
            f"- Never return fewer or more lines.\n"
            
            f"{source_lang} source ({len(lines)} lines):\n"
            f"{numbered_text}\n\n"
            
            f"{dest_lang} translation:\n"
            f"{translation}\n\n"
            
            f"Expected output:\n"
            f"Split the {dest_lang} text into {len(lines)} lines.\n"
            f"Prioritize exact line alignment with the {source_lang} source.\n"
            f"Only apply subtitle readability rules when strict alignment is impossible.\n\n"
            
            f"Example:\n\n"
            
            f"English source:\n"
            f"I really enjoyed the movie\n"
            f"and the amazing soundtrack.\n\n"
            
            f"French translation:\n"
            f"J'ai vraiment apprécié le film et la bande-son incroyable.\n\n"
            
            f"Correct 2-line output:\n"
            f"J'ai vraiment apprécié le film\n"
            f"et la bande-son incroyable."
        )

        # ─────────────────────────────────────────────
        # Persian-specific rules
        # ─────────────────────────────────────────────
        #if dest_lang.lower() == "persian":
        #    prompt = (
        #        f"<ID>\n"
        #    )
        # ─────────────────────────────────────────────
        # Text payload
        # ─────────────────────────────────────────────

        return prompt

    def split_phrase(self, source_lang, dest_lang, source_text, translation):
        # Auto-create doc ID & filename if missing
        if not self.doc_id:
            print("[INFO] No document context. Generating new doc_id and using filename 'inline_text'.")
            #self.set_filename("inline_text")

        prompt = self.build_subtitle_splitter_prompt(source_lang, dest_lang, source_text, translation)
        print("prompt:")
        print(prompt)
        #input("wait")

        input_tokens = self.estimate_tokens(prompt)
        print(f"Estimated number of input tokens: {input_tokens}")

        start_time = time.time()
        if "pro" in self.model:
            response = self.client.responses.create(
                model=self.model,
                input=[
                    {"role": "system", "content": "You are a professional subtitle translator line splitting assistant."},
                    {"role": "user", "content": prompt}
                ]
            )
        else:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a professional subtitle translator line splitting assistant."},
                    {"role": "user", "content": prompt}
                ]
            )
        elapsed_time = time.time() - start_time
        response_json = response.model_dump()

        print("response:")
        print(json.dumps(response_json, indent=4))
        print("--end of response--")

        cost_info = self.calculate_openai_cost(response_json)

        splitted_text = response.choices[0].message.content.strip()
        
        # Remove duplicate new lines if any
        splitted_text = re.sub(r'\n+', '\n', splitted_text)
        print(splitted_text)

        # Validate line counts
        in_lines = source_text.split("\n")
        out_lines = splitted_text.split("\n")
        num_in_lines = 0
        num_newlines = 0
        num_in_lines = source_text.count("\n") + 1
        num_newlines = splitted_text.count("\n") + 1

        if num_in_lines != num_newlines:
            print("[WARNING] Line count mismatch!")
            print(f"Input lines: {len(in_lines)}, Output lines: {len(out_lines)}")
            out_lines = "\n".join(out_lines)
            if len(out_lines) > len(in_lines):
                print("Error in openai line splitting, too many lines")
            else:
                print("Error in openai line splitting, too few lines")
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
                    "split",
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

        return response, out_lines