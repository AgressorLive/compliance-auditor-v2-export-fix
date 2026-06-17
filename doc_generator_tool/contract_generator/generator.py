import json
from .prompts import SYSTEM_PROMPT
from .validator import protect_hallucinations

class ContractGenerator:
    def __init__(self, openai_client=None): self.client = openai_client
    def generate_full_agreement(self, data):
        if not self.client: return self._fallback(data)
        try:
            res = self.client.chat.completions.create(
                model=data.llm_settings.model,
                messages=[{"role":"system","content":SYSTEM_PROMPT},{"role":"user","content":str(data.to_dict())}],
                response_format={"type":"json_object"}, temperature=0.2
            )
            ret = json.loads(res.choices[0].message.content)
            ret["agreement_markdown"] = protect_hallucinations(ret["agreement_markdown"], data)
            return ret
        except: return self._fallback(data)
    def regenerate_section(self, data, section, current):
        return "Regenerated " + section
    def _fallback(self, data):
        from .renderer import render_local_template
        return {"agreement_markdown": render_local_template(data), "missing_required_fields":[], "validation_warnings":[], "generated_defaults":[], "document_status":""}
