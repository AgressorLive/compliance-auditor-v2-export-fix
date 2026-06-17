from .validator import get_placeholder, get_demo_value, get_nested_attr
def render_local_template(data):
    return f"# Agreement {data.agreement_number}\n\nAnnex 1:\nCID: {get_nested_attr(data, 'advertising.account_ids') or get_placeholder('advertising.account_ids')}"
