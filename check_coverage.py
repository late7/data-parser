import json

# Lataa data
with open('consolidated_due_diligence.json', encoding='utf-8') as f:
    data = json.load(f)

# Laske tietopisteet
count = 0
for section_key, section_data in data.items():
    if section_key == 'document_metadata':
        continue

    if isinstance(section_data, dict):
        for sub_key, items in section_data.items():
            if isinstance(items, list):
                count += len(items)
                print(f"  {section_key}.{sub_key}: {len(items)} tietoa")
            else:
                count += 1
                print(f"  {section_key}.{sub_key}: 1 tieto")

print(f"\nYhteensa {count} tietopistetta poimittu")

# Näytä osiot joista ei löytynyt dataa
template_sections = [
    '1_executive_summary',
    '2_company_overview',
    '3_market_and_go_to_market',
    '4_product_and_technology',
    '5_intellectual_property',
    '6_team_and_organization',
    '7_customers_and_traction',
    '8_operations_and_compliance',
    '9_financials',
    '10_funding_and_capital_structure',
    '11_risks_and_dependencies',
    '12_milestones_and_value_creation',
    '13_exit_considerations',
    '14_open_questions_and_gaps',
    '15_appendices'
]

print("\nOsiot joista ei loytynyt dataa:")
for section in template_sections:
    if section not in data:
        print(f"  - {section}")
