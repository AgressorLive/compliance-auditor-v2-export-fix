import json
from .prompts import SYSTEM_PROMPT
from .validator import protect_hallucinations

class ContractGenerator:
    def __init__(self, openai_client=None):
        self.client = openai_client

    def generate_full_agreement(self, data):
        prompt = f"Generate a complete English-language advertising services agreement as JSON based on this data: {json.dumps(data.to_dict())}. Mandatory sections: Title, Parties, Recitals, Scope, Fees, VAT, Accounts, IP, Data Protection, Policies, Notices, Assignment, Severability, Electronic Signatures, Liability, Termination, Governing Law, Force Majeure, Annex 1, Signature Blocks."
        if self.client:
            try:
                res = self.client.chat.completions.create(
                    model=data.llm_settings.model,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": prompt}
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.2
                )
                ret = json.loads(res.choices[0].message.content)
                ret["agreement_markdown"] = protect_hallucinations(ret.get("agreement_markdown", ""), data)
                return ret
            except Exception:
                return self._fallback(data)
        return self._fallback(data)

    def regenerate_section(self, data, section_name, current_agreement):
        if not self.client:
            return "AI Generation unavailable."

        prompt = (
            f"Regenerate the section '{section_name}' for this agreement. "
            f"Verified data: {json.dumps(data.to_dict())}. "
            f"Keep all existing verified values exactly as they are. "
            f"Current context:\n\n{current_agreement}"
        )
        try:
            res = self.client.chat.completions.create(
                model=data.llm_settings.model,
                messages=[
                    {"role": "system", "content": "You are a legal document editor. Output ONLY the new section text, no commentary."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2
            )
            return protect_hallucinations(res.choices[0].message.content, data)
        except Exception as e:
            return f"Error: {str(e)}"

    def _fallback(self, data):
        from .renderer import render_local_template
        return {
            "agreement_markdown": render_local_template(data),
            "missing_required_fields": [],
            "validation_warnings": ["Local template fallback used."],
            "generated_defaults": [],
            "document_status": "missing_critical_fields"
        }
