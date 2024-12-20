import deepl 
auth_key="4ec6a0ba-b098-1c54-1c6f-d61b777b71f4:fx"
translator = deepl.Translator(auth_key) 
result = translator.translate_text("hello", target_lang="fr") 
translated_text = result.text
print(translated_text)