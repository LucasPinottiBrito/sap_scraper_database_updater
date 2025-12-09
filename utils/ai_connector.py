import base64
import re
from dotenv import load_dotenv
import os
import json
from utils.prompts import document_analisys_prompt
from openai import OpenAI
load_dotenv()

databricks_url = os.getenv("DATABRICKS_URL")
databricks_token = os.getenv("DATABRICKS_TOKEN")

client = OpenAI(
    api_key=databricks_token,
    base_url=f"{databricks_url}/serving-endpoints/"
)

class AIConnector:
    def __init__(self):
        self.ai_model = "databricks-claude-3-7-sonnet"
        self.url = f"{databricks_url}/serving-endpoints/{self.ai_model}/invocations"
        self.headers = {
            "Authorization": f"Bearer {databricks_token}",
            "Content-Type": "application/json",
        }
    
    def _flatten_json(self, d, parent_key='', sep='_'):
        items = {}
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.update(self._flatten_json(v, new_key, sep=sep))
            else:
                items[new_key] = v
        return items

    def _extract_json(self, result_text):
        try:
            result_text = re.sub(r"```.*?(\r\n|\n)?", "", result_text)
            json_match = re.search(r"\{.*\}", result_text, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
        except Exception as e:
            print(f"Erro ao extrair JSON: {e}")
        return {}

    def get_ai_response(self, prompt: str) -> str:
        completion = client.chat.completions.create(
            model=self.ai_model,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=100,
            temperature=0.7
        )
        return completion.choices[0].message.content

    def get_pdf_b64_analysis(self, pdf_path_list: list) -> str:
        try:
            content_array = [
                {
                    "type": "text",
                    "text": f"{document_analisys_prompt}"
                }
            ]
            for path in pdf_path_list:
                with open(path, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                    pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
                    content_array.append(
                    {
                        "type": "document",
                        "source": {"url": f"data:application/pdf;base64,{pdf_b64}"}
                    })
        except Exception as e:
            print(f"Erro ao ler ou codificar o PDF: {e}")
            return {}
        try:
            completion = client.chat.completions.create(
                model="databricks-claude-3-7-sonnet",
                messages=[{"role": "user", "content": content_array}],
                max_tokens=400 # Boa prática definir um max_tokens
            )
            return self._extract_json(completion.choices[0].message.content)
        except Exception as e:
            print(f"Erro ao obter resposta da IA: {e}")
            return {}