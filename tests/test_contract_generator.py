import unittest, sys, os, re
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'doc_generator_tool')))
from contract_generator.models import AgreementData, PartyDetails
from contract_generator.validator import validate_agreement, get_placeholder, protect_hallucinations

class TestContractGenerator(unittest.TestCase):
    def test_production_mode_no_vat(self):
        data = AgreementData(mode='production')
        missing, warns = validate_agreement(data, set())
        self.assertIn('contractor.vat_number', missing)
    def test_hallucination_protection(self):
        data = AgreementData(mode='production', contractor=PartyDetails(iban='EE123'))
        text = 'IBAN is EE123 and fake is GB99ABCD123456789012.'
        protected = protect_hallucinations(text, data)
        self.assertIn('[REQUIRED: IBAN]', protected)
        self.assertNotIn('GB99ABCD', protected)
if __name__ == '__main__':
    unittest.main()
