from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import List, Optional, Dict, Any

class AgreementMode(Enum):
    PRODUCTION = "production"
    DEMO = "demo"

class FieldState(Enum):
    MISSING = "missing"
    USER_PROVIDED = "user_provided"
    GENERATED_DEFAULT = "generated_default"
    VERIFIED = "verified"
    DEMO_FICTIONAL = "demo_fictional"

@dataclass
class PartyDetails:
    trade_name: Optional[str] = None
    legal_name: Optional[str] = None
    legal_form: Optional[str] = None
    country: Optional[str] = None
    registration_number: Optional[str] = None
    registered_address: Optional[str] = None
    vat_number: Optional[str] = None
    vat_status: Optional[str] = None
    signatory_name: Optional[str] = None
    signatory_position: Optional[str] = None
    authority_basis: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    bank_account_holder: Optional[str] = None
    iban: Optional[str] = None
    swift_bic: Optional[str] = None
    bank_name: Optional[str] = None

@dataclass
class AdvertisingDetails:
    platform: str = "Google Ads"
    account_owner: Optional[str] = None
    billing_party: Optional[str] = None
    account_ids: List[str] = field(default_factory=list)
    media_budget_included: bool = False
    media_budget_amount: Optional[float] = None
    services: List[str] = field(default_factory=list)
    deliverables: List[str] = field(default_factory=list)
    target_countries: List[str] = field(default_factory=list)

@dataclass
class LegalOptions:
    governing_law: Optional[str] = None
    court_city: Optional[str] = None
    court_country: Optional[str] = None
    late_fee_percent_per_day: float = 0.1
    late_fee_cap_percent: float = 10.0
    access_return_days: int = 5
    breach_cure_days: int = 10
    liability_cap: str = "fees paid under the Agreement"
    ip_model: str = "assignment_on_full_payment"
    notice_email_allowed: bool = True

@dataclass
class LLMSettings:
    model: str = "gpt-4o"
    timeout: int = 120
    base_url: Optional[str] = None
    temperature: float = 0.2

@dataclass
class AgreementData:
    mode: str = "production"
    agreement_number: Optional[str] = None
    agreement_date: Optional[str] = None
    service_fee: float = 0.0
    currency: str = "EUR"
    contractor: PartyDetails = field(default_factory=PartyDetails)
    client: PartyDetails = field(default_factory=PartyDetails)
    advertising: AdvertisingDetails = field(default_factory=AdvertisingDetails)
    legal_options: LegalOptions = field(default_factory=LegalOptions)
    llm_settings: LLMSettings = field(default_factory=LLMSettings)

    def to_dict(self): return asdict(self)

@dataclass
class ProjectData:
    schema_version: int = 1
    created_at: str = ""
    updated_at: str = ""
    agreement_data: AgreementData = field(default_factory=AgreementData)
    field_states: Dict[str, str] = field(default_factory=dict)
    generated_sections: Dict[str, str] = field(default_factory=dict)
    missing_required_fields: List[str] = field(default_factory=list)
    validation_warnings: List[str] = field(default_factory=list)
    document_status: str = "missing_critical_fields"

    def to_dict(self): return asdict(self)
