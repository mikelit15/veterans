import re

badWorksheet = "Evergreen"
folder_name = "A"
formatted_id = "02241"
orig_target = "Cemetery - Redacted/Fairview - Redacted/Z/FairviewB08223 redacted.pdf"
modified_string = re.sub(r'(/[A-Z]/)', f'/{folder_name}/', orig_target)
modified_string = re.sub(fr'(Fairview[A-Z])\d{{5}}', f'{badWorksheet}{folder_name}{formatted_id}', modified_string)

print(modified_string)